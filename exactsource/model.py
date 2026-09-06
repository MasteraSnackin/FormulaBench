"""Small, fixed Tinker client used by the ExactSource runtime."""

from __future__ import annotations

import hashlib
import json
import math
import os
import re
import time
from codecs import getincrementaldecoder
from collections.abc import Callable, Iterator, Mapping, Sequence
from dataclasses import dataclass, field
from datetime import UTC, datetime
from email.utils import parsedate_to_datetime
from typing import Any

import httpx

from exactsource.config import (
    API_KEY_ENV,
    MAX_OUTPUT_TOKENS,
    MAX_RETRY_AFTER_SECONDS,
    MAX_RETRY_STAGGER_SECONDS,
    MODEL_CONNECT_TIMEOUT_SECONDS,
    MODEL_ID,
    MODEL_POOL_TIMEOUT_SECONDS,
    MODEL_STREAM_READ_TIMEOUT_SECONDS,
    MODEL_WRITE_TIMEOUT_SECONDS,
    REASONING_EFFORT,
    TEMPERATURE,
    TINKER_MESSAGES_URL,
    TRANSPORT_RETRIES,
)
from exactsource.contracts import ModelReply

_RETRYABLE_STATUS = frozenset({408, 409, 425, 429, 500, 502, 503, 504, 529})
_RETRYABLE_STREAM_ERRORS = frozenset({"overloaded_error"})
_SECRET_PATTERNS = (
    re.compile(r"(?i)(authorization\s*[:=]\s*bearer\s+)[^\s,;]+"),
    re.compile(r"(?i)(x-api-key\s*[:=]\s*)[^\s,;]+"),
    re.compile(r"(?i)(bearer\s+)[A-Za-z0-9._~+/=-]{8,}"),
    re.compile(r"\bsk-or-v1-[A-Za-z0-9_-]+\b"),
)
_LEGACY_FENCED_JSON_OBJECT_RE = re.compile(
    r"^```(?:json)?\s*(\{.*\})\s*```$",
    re.IGNORECASE | re.DOTALL,
)


class ModelError(RuntimeError):
    """Base exception for reproducible model-call failures."""


class ModelConfigurationError(ModelError):
    """Raised when required local credentials are absent."""


class ModelTransportError(ModelError):
    """Raised when Tinker cannot return a successful response."""


class ModelResponseError(ModelError):
    """Raised when a successful response violates the expected response shape."""


class ModelTruncationError(ModelResponseError):
    """Raised when the provider reaches its output-token cap before completion."""

    def __init__(self, message: str, *, output_tokens: int | None) -> None:
        super().__init__(message)
        self.output_tokens = output_tokens
        self.stop_reason = "max_tokens"


class _ProviderStreamError(ModelTransportError):
    """Structured error emitted inside an otherwise successful SSE response."""

    def __init__(self, error_type: str, message: str) -> None:
        super().__init__(f"Tinker stream error {error_type}: {message}")
        self.error_type = error_type


def redact_secrets(value: object, secrets: Sequence[str] = ()) -> str:
    """Return a log-safe rendering with known and recognisable keys removed."""

    text = str(value)
    env_secret = os.environ.get(API_KEY_ENV)
    candidates = [*secrets]
    if env_secret:
        candidates.append(env_secret)
    for secret in sorted({item for item in candidates if item}, key=len, reverse=True):
        text = text.replace(secret, "[REDACTED]")
    for pattern in _SECRET_PATTERNS:
        text = pattern.sub(
            lambda match: f"{match.group(1)}[REDACTED]" if match.lastindex else "[REDACTED]",
            text,
        )
    return text


def _validated_messages(messages: list[dict[str, str]]) -> list[dict[str, str]]:
    if not messages:
        raise ValueError("messages must not be empty")
    allowed_roles = {"system", "user", "assistant"}
    validated: list[dict[str, str]] = []
    for index, message in enumerate(messages):
        if set(message) != {"role", "content"}:
            raise ValueError(f"message {index} must contain only role and content")
        role = message["role"]
        content = message["content"]
        if role not in allowed_roles:
            raise ValueError(f"message {index} has unsupported role {role!r}")
        if not isinstance(content, str) or not content.strip():
            raise ValueError(f"message {index} content must be a non-empty string")
        validated.append({"role": role, "content": content})
    return validated


def _required_token_count(usage: Mapping[str, object], name: str) -> int:
    value = usage.get(name)
    if isinstance(value, bool) or not isinstance(value, int) or value < 0:
        raise ModelResponseError(f"provider response usage has invalid {name}")
    return value


def _optional_token_count(usage: Mapping[str, object], name: str) -> int:
    value = usage.get(name, 0)
    if value is None:
        return 0
    if isinstance(value, bool) or not isinstance(value, int) or value < 0:
        raise ModelResponseError(f"provider response usage has invalid {name}")
    return value


def _anthropic_messages(
    messages: list[dict[str, str]],
) -> tuple[str | None, list[dict[str, str]]]:
    """Move system instructions to Anthropic's top-level ``system`` field."""

    system_parts: list[str] = []
    conversation: list[dict[str, str]] = []
    for message in messages:
        if message["role"] == "system":
            system_parts.append(message["content"])
        else:
            conversation.append(message)
    if not conversation:
        raise ValueError("messages must contain at least one user or assistant message")
    return ("\n\n".join(system_parts) if system_parts else None, conversation)


def _validated_max_output_tokens(value: object) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value <= 0:
        raise ValueError("max_output_tokens must be a positive integer")
    return value


def _validated_reasoning_effort(value: object) -> bool:
    if not isinstance(value, bool):
        raise ValueError("reasoning_effort must be a boolean")
    return value


def serialise_request_payload(
    messages: list[dict[str, str]],
    *,
    max_output_tokens: int = MAX_OUTPUT_TOKENS,
    reasoning_effort: bool = REASONING_EFFORT,
) -> tuple[list[dict[str, str]], dict[str, Any], str]:
    """Build one validated provider payload without making a provider request."""

    validated = _validated_messages(messages)
    effective_max_output_tokens = _validated_max_output_tokens(max_output_tokens)
    effective_reasoning_effort = _validated_reasoning_effort(reasoning_effort)
    system, conversation = _anthropic_messages(validated)
    payload: dict[str, Any] = {
        "model": MODEL_ID,
        "messages": conversation,
        "temperature": TEMPERATURE,
        "max_tokens": effective_max_output_tokens,
        "reasoning_effort": effective_reasoning_effort,
        "stream": True,
    }
    if system is not None:
        payload["system"] = system
    serialised = json.dumps(
        payload,
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
    )
    return validated, payload, serialised


def _split_legacy_thinking(text: str) -> tuple[str, str]:
    """Split a legacy text-only reasoning prelude from its JSON-object answer.

    Compatible endpoints can place reasoning and the final answer in one text
    block separated by ``</think>``. Reasoning may itself contain draft JSON, and
    the final JSON may contain the delimiter as string data, so delimiter
    candidates are checked from right to left. A split is accepted only when the
    complete trimmed suffix is a JSON object. Otherwise the original trimmed text
    remains the answer and no reasoning is inferred.
    """

    candidate = text.strip()
    closing = "</think>"
    marker = candidate.rfind(closing)
    while marker >= 0:
        answer = candidate[marker + len(closing) :].strip()
        fenced = _LEGACY_FENCED_JSON_OBJECT_RE.fullmatch(answer)
        json_candidate = fenced.group(1).strip() if fenced else answer
        try:
            decoded = json.loads(json_candidate)
        except json.JSONDecodeError:
            pass
        else:
            if isinstance(decoded, dict):
                reasoning = candidate[:marker]
                if reasoning.startswith("<think>"):
                    reasoning = reasoning[len("<think>") :]
                return answer, reasoning
        marker = candidate.rfind(closing, 0, marker)
    return candidate, ""


def _retry_stagger(payload_fingerprint: bytes, attempt: int) -> float:
    """Spread concurrent task retries reproducibly without global random state."""

    digest = hashlib.sha256(payload_fingerprint + attempt.to_bytes(2, "big")).digest()
    fraction = int.from_bytes(digest[:2], "big") / 65_535
    return fraction * MAX_RETRY_STAGGER_SECONDS


def _retry_delay(
    response: httpx.Response | None,
    attempt: int,
    payload_fingerprint: bytes,
) -> float:
    """Return a bounded Retry-After delay, with deterministic backoff fallback."""

    stagger = _retry_stagger(payload_fingerprint, attempt)
    fallback = min(0.5 * (2**attempt) + stagger, MAX_RETRY_AFTER_SECONDS)
    value = response.headers.get("retry-after") if response is not None else None
    if value is None:
        return fallback

    candidate: float | None = None
    try:
        candidate = float(value.strip())
    except ValueError:
        try:
            retry_at = parsedate_to_datetime(value)
            if retry_at.tzinfo is None:
                retry_at = retry_at.replace(tzinfo=UTC)
            candidate = (retry_at - datetime.now(UTC)).total_seconds()
        except (TypeError, ValueError, OverflowError):
            return fallback
    if not math.isfinite(candidate) or candidate < 0:
        return fallback
    return min(candidate + stagger, MAX_RETRY_AFTER_SECONDS)


def _request_id(response: httpx.Response) -> str | None:
    return response.headers.get("request-id") or response.headers.get("x-request-id")


def _response_preview(response: httpx.Response) -> str:
    """Read a small error-body preview without masking the HTTP status itself."""

    try:
        response.read()
        return response.text[:2_000]
    except httpx.HTTPError as exc:
        return f"<response body unavailable: {exc}>"


def _parse_sse_events(
    response: httpx.Response,
    progress: _StreamProgress,
) -> Iterator[tuple[str, Mapping[str, object]]]:
    """Parse Anthropic's event/data SSE frames and reject ambiguous framing."""

    event_name: str | None = None
    data_lines: list[str] = []
    has_fields = False

    for line in _stream_lines(response, progress):
        if line == "":
            if not has_fields:
                continue
            if event_name is None:
                raise ModelResponseError("provider SSE frame is missing its event field")
            if not data_lines:
                raise ModelResponseError("provider SSE frame is missing its data field")
            try:
                payload = json.loads("\n".join(data_lines))
            except json.JSONDecodeError:
                raise ModelResponseError("provider SSE frame contains invalid JSON") from None
            if not isinstance(payload, Mapping):
                raise ModelResponseError("provider SSE data must be a JSON object")
            payload_type = payload.get("type")
            if payload_type != event_name:
                raise ModelResponseError("provider SSE event field does not match its JSON type")
            yield event_name, payload
            event_name = None
            data_lines = []
            has_fields = False
            continue

        if line.startswith(":"):
            continue
        if ":" in line:
            field_name, value = line.split(":", 1)
            if value.startswith(" "):
                value = value[1:]
        else:
            field_name, value = line, ""

        if field_name == "event":
            if event_name is not None:
                raise ModelResponseError("provider SSE frame repeats its event field")
            if not value:
                raise ModelResponseError("provider SSE frame has an empty event field")
            event_name = value
            has_fields = True
        elif field_name == "data":
            data_lines.append(value)
            has_fields = True
        else:
            raise ModelResponseError(f"provider SSE frame has unsupported field {field_name!r}")

    if has_fields:
        raise ModelResponseError("provider SSE stream ended in the middle of an event")


def _stream_lines(response: httpx.Response, progress: _StreamProgress) -> Iterator[str]:
    """Yield UTF-8 lines while recording the first raw byte before buffering it."""

    decoder = getincrementaldecoder("utf-8")(errors="strict")
    buffer = ""

    def complete_lines(*, final: bool) -> Iterator[str]:
        nonlocal buffer
        while True:
            lf_index = buffer.find("\n")
            cr_index = buffer.find("\r")
            indexes = [index for index in (lf_index, cr_index) if index >= 0]
            if not indexes:
                break
            index = min(indexes)
            if buffer[index] == "\r" and index == len(buffer) - 1 and not final:
                break
            terminator_length = (
                2
                if buffer[index] == "\r" and index + 1 < len(buffer) and buffer[index + 1] == "\n"
                else 1
            )
            yield buffer[:index]
            buffer = buffer[index + terminator_length :]

    try:
        for chunk in response.iter_bytes():
            if not chunk:
                continue
            downloaded = response.num_bytes_downloaded
            if downloaded > progress.raw_bytes_received:
                progress.raw_bytes_received = downloaded
            else:
                # Directly constructed/cached httpx responses do not always update
                # num_bytes_downloaded; retain a deterministic decoded-byte fallback.
                progress.raw_bytes_received += len(chunk)
            buffer += decoder.decode(chunk, final=False)
            yield from complete_lines(final=False)
        buffer += decoder.decode(b"", final=True)
    except UnicodeDecodeError:
        raise ModelResponseError("provider SSE stream is not valid UTF-8") from None

    yield from complete_lines(final=True)
    if buffer:
        yield buffer


def _event_index(payload: Mapping[str, object]) -> int:
    index = payload.get("index")
    if isinstance(index, bool) or not isinstance(index, int) or index < 0:
        raise ModelResponseError("provider content event has an invalid block index")
    return index


@dataclass(slots=True)
class _ContentBlock:
    kind: str
    parts: list[str] = field(default_factory=list)
    closed: bool = False


@dataclass(slots=True)
class _StreamProgress:
    raw_bytes_received: int = 0


class _AnthropicStreamDecoder:
    """Stateful, strict decoder for the Anthropic Messages streaming protocol."""

    def __init__(self) -> None:
        self.started = False
        self.message_delta_count = 0
        self.stopped = False
        self.generated_content = False
        self.active_index: int | None = None
        self.blocks: list[_ContentBlock] = []
        self.input_tokens: int | None = None
        self.output_tokens: int | None = None
        self.cache_creation_input_tokens: int | None = None
        self.cache_read_input_tokens: int | None = None
        self.stop_reason: str | None = None

    def feed(self, event: str, payload: Mapping[str, object]) -> None:
        if event == "ping":
            return
        if self.stopped:
            raise ModelResponseError("provider sent data after message_stop")
        if event == "error":
            error = payload.get("error")
            if not isinstance(error, Mapping):
                raise ModelResponseError("provider SSE error event is malformed")
            error_type = error.get("type")
            message = error.get("message")
            if not isinstance(error_type, str) or not isinstance(message, str):
                raise ModelResponseError("provider SSE error event is malformed")
            raise _ProviderStreamError(error_type, message)
        if event == "message_start":
            self._message_start(payload)
        elif event == "content_block_start":
            self._content_block_start(payload)
        elif event == "content_block_delta":
            self._content_block_delta(payload)
        elif event == "content_block_stop":
            self._content_block_stop(payload)
        elif event == "message_delta":
            self._message_delta(payload)
        elif event == "message_stop":
            self._message_stop()
        else:
            # Anthropic may add new top-level events. Their framing remains
            # validated, but an unknown event must not break an otherwise complete
            # response when it does not mutate the states understood here.
            return

    def _message_start(self, payload: Mapping[str, object]) -> None:
        if self.started:
            raise ModelResponseError("provider sent message_start more than once")
        message = payload.get("message")
        if not isinstance(message, Mapping):
            raise ModelResponseError("provider message_start is missing its message")
        if message.get("type") != "message" or message.get("role") != "assistant":
            raise ModelResponseError("provider message_start has invalid message metadata")
        returned_model = message.get("model")
        if returned_model != MODEL_ID:
            raise ModelResponseError(
                f"provider returned unexpected model identity {returned_model!r}"
            )
        content = message.get("content")
        if content != []:
            raise ModelResponseError("provider message_start content must be empty")
        usage = message.get("usage")
        if not isinstance(usage, Mapping):
            raise ModelResponseError("provider response usage is malformed")
        base_input = _required_token_count(usage, "input_tokens")
        self.cache_creation_input_tokens = _optional_token_count(
            usage, "cache_creation_input_tokens"
        )
        self.cache_read_input_tokens = _optional_token_count(usage, "cache_read_input_tokens")
        self.input_tokens = (
            base_input + self.cache_creation_input_tokens + self.cache_read_input_tokens
        )
        if "output_tokens" in usage:
            self.output_tokens = _required_token_count(usage, "output_tokens")
        self.started = True

    def _require_message_body(self) -> None:
        if not self.started:
            raise ModelResponseError("provider sent content before message_start")
        if self.message_delta_count:
            raise ModelResponseError("provider sent content after message_delta")

    def _content_block_start(self, payload: Mapping[str, object]) -> None:
        self._require_message_body()
        if self.active_index is not None:
            raise ModelResponseError("provider started overlapping content blocks")
        index = _event_index(payload)
        if index != len(self.blocks):
            raise ModelResponseError("provider content block indexes are not contiguous")
        content_block = payload.get("content_block")
        if not isinstance(content_block, Mapping):
            raise ModelResponseError("provider content_block_start is malformed")
        kind = content_block.get("type")
        if kind == "text":
            initial = content_block.get("text")
        elif kind == "thinking":
            initial = content_block.get("thinking")
        else:
            raise ModelResponseError(f"provider content block {index} is not text or thinking")
        if not isinstance(initial, str):
            raise ModelResponseError(f"provider content block {index} has invalid initial content")
        self.blocks.append(_ContentBlock(kind=kind, parts=[initial]))
        if initial:
            self.generated_content = True
        self.active_index = index

    def _content_block_delta(self, payload: Mapping[str, object]) -> None:
        self._require_message_body()
        index = _event_index(payload)
        if self.active_index != index:
            raise ModelResponseError("provider content delta does not match the active block")
        delta = payload.get("delta")
        if not isinstance(delta, Mapping):
            raise ModelResponseError("provider content_block_delta is malformed")
        block = self.blocks[index]
        delta_type = delta.get("type")
        if block.kind == "text" and delta_type == "text_delta":
            value = delta.get("text")
        elif block.kind == "thinking" and delta_type == "thinking_delta":
            value = delta.get("thinking")
        elif block.kind == "thinking" and delta_type == "signature_delta":
            signature = delta.get("signature")
            if not isinstance(signature, str):
                raise ModelResponseError("provider thinking signature is malformed")
            if signature:
                self.generated_content = True
            return
        else:
            raise ModelResponseError(
                f"provider delta type {delta_type!r} does not match {block.kind} block"
            )
        if not isinstance(value, str):
            raise ModelResponseError("provider content delta has invalid text")
        block.parts.append(value)
        if value:
            self.generated_content = True

    def _content_block_stop(self, payload: Mapping[str, object]) -> None:
        self._require_message_body()
        index = _event_index(payload)
        if self.active_index != index:
            raise ModelResponseError("provider stopped a content block that is not active")
        self.blocks[index].closed = True
        self.active_index = None

    def _message_delta(self, payload: Mapping[str, object]) -> None:
        if not self.started:
            raise ModelResponseError("provider sent message_delta before message_start")
        if self.active_index is not None:
            raise ModelResponseError("provider sent message_delta before content block stop")
        if self.stop_reason is not None:
            raise ModelResponseError("provider sent message_delta after terminal stop_reason")
        delta = payload.get("delta")
        if not isinstance(delta, Mapping):
            raise ModelResponseError("provider message_delta is malformed")
        stop_reason = delta.get("stop_reason")
        if stop_reason is not None and (not isinstance(stop_reason, str) or not stop_reason):
            raise ModelResponseError("provider message_delta has invalid stop_reason")
        usage = payload.get("usage")
        if not isinstance(usage, Mapping):
            raise ModelResponseError("provider response usage is malformed")
        output_tokens = _required_token_count(usage, "output_tokens")
        if self.output_tokens is not None and output_tokens < self.output_tokens:
            raise ModelResponseError("provider cumulative output token usage decreased")
        self.output_tokens = output_tokens
        self.stop_reason = stop_reason
        self.message_delta_count += 1
        if stop_reason == "max_tokens":
            raise ModelTruncationError(
                "provider stopped before a clean end_turn ('max_tokens')",
                output_tokens=output_tokens,
            )
        if stop_reason is not None and stop_reason != "end_turn":
            raise ModelResponseError(f"provider stopped before a clean end_turn ({stop_reason!r})")

    def _message_stop(self) -> None:
        if not self.message_delta_count:
            raise ModelResponseError("provider sent message_stop before message_delta")
        if self.stop_reason != "end_turn":
            raise ModelResponseError("provider message_stop has no terminal end_turn")
        self.stopped = True

    def raw_text(self) -> str:
        parts: list[str] = []
        for block in self.blocks:
            content = "".join(block.parts)
            if block.kind == "thinking":
                parts.append(f"<think>{content}")
                if block.closed:
                    parts.append("</think>")
            else:
                parts.append(content)
        return "".join(parts)

    def answer_text(self) -> str:
        """Return generated answer blocks without confusing answer data for reasoning."""

        text = "".join("".join(block.parts) for block in self.blocks if block.kind == "text")
        if any(block.kind == "thinking" for block in self.blocks):
            answer = text.strip()
            if not answer:
                raise ModelResponseError("provider response has no answer after reasoning")
            return answer
        answer, _ = _split_legacy_thinking(text)
        return answer

    def thinking_text(self) -> str:
        """Return reasoning content without synthetic ``<think>`` wrappers."""

        native = "".join("".join(block.parts) for block in self.blocks if block.kind == "thinking")
        if any(block.kind == "thinking" for block in self.blocks):
            return native
        text = "".join("".join(block.parts) for block in self.blocks if block.kind == "text")
        _, thinking = _split_legacy_thinking(text)
        return thinking

    def finish(self) -> dict[str, Any]:
        if not self.stopped:
            raise ModelResponseError("provider SSE stream ended before message_stop")
        if not self.blocks:
            raise ModelResponseError("provider response content must be a non-empty list")
        if (
            self.input_tokens is None
            or self.output_tokens is None
            or self.cache_creation_input_tokens is None
            or self.cache_read_input_tokens is None
        ):
            raise ModelResponseError("provider response usage is incomplete")
        raw_text = self.raw_text()
        if not raw_text.strip():
            raise ModelResponseError("provider response content is empty")
        return {
            "text": self.answer_text(),
            "raw_text": raw_text,
            "input_tokens": self.input_tokens,
            "output_tokens": self.output_tokens,
            "cache_creation_input_tokens": self.cache_creation_input_tokens,
            "cache_read_input_tokens": self.cache_read_input_tokens,
            "stop_reason": self.stop_reason,
        }


def _decode_until_message_stop(
    response: httpx.Response,
    progress: _StreamProgress,
    decoder: _AnthropicStreamDecoder,
) -> dict[str, Any]:
    """Return at the protocol completion boundary without draining the socket."""

    for event, event_payload in _parse_sse_events(response, progress):
        decoder.feed(event, event_payload)
        if decoder.stopped:
            return decoder.finish()
    return decoder.finish()


class TinkerClient:
    """Synchronous streaming client with a fixed model and bounded retries.

    An ``httpx.Client`` and sleeper can be injected for deterministic tests. The API
    key is never included in ``repr`` or exception messages.
    """

    def __init__(
        self,
        *,
        api_key: str | None = None,
        client: httpx.Client | None = None,
        sleep: Callable[[float], None] = time.sleep,
    ) -> None:
        key = api_key if api_key is not None else os.environ.get(API_KEY_ENV)
        if not key or not key.strip():
            raise ModelConfigurationError(
                f"{API_KEY_ENV} is required; configure it in the execution environment"
            )
        self._api_key = key.strip()
        self._timeout = httpx.Timeout(
            connect=MODEL_CONNECT_TIMEOUT_SECONDS,
            read=MODEL_STREAM_READ_TIMEOUT_SECONDS,
            write=MODEL_WRITE_TIMEOUT_SECONDS,
            pool=MODEL_POOL_TIMEOUT_SECONDS,
        )
        self._client = client or httpx.Client(timeout=self._timeout)
        self._owns_client = client is None
        self._sleep = sleep

    def __repr__(self) -> str:
        return f"TinkerClient(model={MODEL_ID!r}, api_key='[REDACTED]')"

    def close(self) -> None:
        if self._owns_client:
            self._client.close()

    def __enter__(self) -> TinkerClient:
        return self

    def __exit__(self, *_: object) -> None:
        self.close()

    def complete(
        self,
        messages: list[dict[str, str]],
        *,
        max_output_tokens: int = MAX_OUTPUT_TOKENS,
        reasoning_effort: bool = REASONING_EFFORT,
        on_attempt: Callable[[dict[str, object]], None] | None = None,
    ) -> ModelReply:
        """Stream a completion, reporting every transport attempt when requested.

        The callback receives only trace-safe values. It fires once for every HTTP
        attempt. A request is retried only before generated content arrives, or when
        the server returns an explicitly retryable HTTP status. Partial generations
        are never replayed because doing so can duplicate paid inference.
        """

        validated, payload, serialised_payload = serialise_request_payload(
            messages,
            max_output_tokens=max_output_tokens,
            reasoning_effort=reasoning_effort,
        )
        effective_max_output_tokens = _validated_max_output_tokens(max_output_tokens)
        effective_reasoning_effort = _validated_reasoning_effort(reasoning_effort)
        payload_fingerprint = serialised_payload.encode()
        request_sha256 = hashlib.sha256(payload_fingerprint).hexdigest()
        request_chars = len(serialised_payload)
        message_chars = sum(len(message["content"]) for message in validated)
        headers = {
            "x-api-key": self._api_key,
            "anthropic-version": "2023-06-01",
            "Accept": "text/event-stream",
            "Content-Type": "application/json",
        }

        def report_attempt(**fields: Any) -> None:
            self._report_attempt(
                on_attempt,
                max_output_tokens=effective_max_output_tokens,
                reasoning_effort=effective_reasoning_effort,
                request_sha256=request_sha256,
                **fields,
            )

        started = time.monotonic()
        last_problem = "unknown transport failure"
        for attempt in range(TRANSPORT_RETRIES + 1):
            attempt_started = time.monotonic()
            progress = _StreamProgress()
            decoder = _AnthropicStreamDecoder()
            response_status: int | None = None
            response_request_id: str | None = None
            active_response: httpx.Response | None = None
            try:
                with self._client.stream(
                    "POST",
                    TINKER_MESSAGES_URL,
                    headers=headers,
                    json=payload,
                    timeout=self._timeout,
                ) as response:
                    active_response = response
                    response_status = response.status_code
                    response_request_id = _request_id(response)
                    if response.status_code in _RETRYABLE_STATUS:
                        preview = _response_preview(response)
                        progress.raw_bytes_received = max(
                            progress.raw_bytes_received,
                            response.num_bytes_downloaded,
                            len(preview.encode()),
                        )
                        detail = redact_secrets(preview[:500], (self._api_key,))
                        last_problem = f"HTTP {response.status_code}: {detail}"
                        retry_delay = _retry_delay(response, attempt, payload_fingerprint)
                        report_attempt(
                            attempt=attempt + 1,
                            status=("retry" if attempt < TRANSPORT_RETRIES else "error"),
                            retryable=True,
                            http_status=response.status_code,
                            response_text=preview,
                            error=last_problem,
                            started=attempt_started,
                            retry_delay_seconds=(
                                retry_delay if attempt < TRANSPORT_RETRIES else None
                            ),
                            request_id=response_request_id,
                            raw_bytes_received=progress.raw_bytes_received,
                            request_chars=request_chars,
                            message_chars=message_chars,
                        )
                        if attempt >= TRANSPORT_RETRIES:
                            raise ModelTransportError(
                                "Tinker transport failed after "
                                f"{attempt + 1} attempts: {last_problem}"
                            )
                        response.close()
                        self._sleep(retry_delay)
                        continue

                    if response.status_code != 200:
                        preview = _response_preview(response)
                        progress.raw_bytes_received = max(
                            progress.raw_bytes_received,
                            response.num_bytes_downloaded,
                            len(preview.encode()),
                        )
                        problem = redact_secrets(preview[:500], (self._api_key,))
                        status = response.status_code
                        report_attempt(
                            attempt=attempt + 1,
                            status="error",
                            retryable=False,
                            http_status=status,
                            response_text=preview,
                            error=problem,
                            started=attempt_started,
                            request_id=response_request_id,
                            raw_bytes_received=progress.raw_bytes_received,
                            request_chars=request_chars,
                            message_chars=message_chars,
                        )
                        raise ModelTransportError(
                            f"Tinker rejected the request with HTTP {status}: {problem}"
                        )

                    media_type = (
                        response.headers.get("content-type", "").split(";", 1)[0].strip().casefold()
                    )
                    if media_type != "text/event-stream":
                        preview = _response_preview(response)
                        error = "provider returned a non-SSE success response"
                        report_attempt(
                            attempt=attempt + 1,
                            status="error",
                            retryable=False,
                            http_status=response.status_code,
                            response_text=preview,
                            error=error,
                            started=attempt_started,
                            request_id=response_request_id,
                            raw_bytes_received=max(
                                response.num_bytes_downloaded,
                                len(preview.encode()),
                            ),
                            request_chars=request_chars,
                            message_chars=message_chars,
                        )
                        raise ModelResponseError(error)

                    try:
                        reply = _decode_until_message_stop(response, progress, decoder)
                    except _ProviderStreamError as exc:
                        problem = redact_secrets(exc, (self._api_key,))
                        retryable = (
                            exc.error_type in _RETRYABLE_STREAM_ERRORS
                            and not decoder.generated_content
                        )
                        retry_delay = _retry_delay(
                            response,
                            attempt,
                            payload_fingerprint,
                        )
                        report_attempt(
                            attempt=attempt + 1,
                            status=(
                                "retry" if retryable and attempt < TRANSPORT_RETRIES else "error"
                            ),
                            retryable=retryable,
                            http_status=response.status_code,
                            response_text=decoder.raw_text(),
                            error=problem,
                            started=attempt_started,
                            input_tokens=decoder.input_tokens,
                            output_tokens=decoder.output_tokens,
                            cache_creation_input_tokens=(decoder.cache_creation_input_tokens),
                            cache_read_input_tokens=decoder.cache_read_input_tokens,
                            retry_delay_seconds=(
                                retry_delay if retryable and attempt < TRANSPORT_RETRIES else None
                            ),
                            stop_reason=decoder.stop_reason,
                            request_id=response_request_id,
                            raw_bytes_received=progress.raw_bytes_received,
                            generated_content=decoder.generated_content,
                            message_complete=decoder.stopped,
                            thinking_text=decoder.thinking_text(),
                            request_chars=request_chars,
                            message_chars=message_chars,
                        )
                        if retryable and attempt < TRANSPORT_RETRIES:
                            response.close()
                            self._sleep(retry_delay)
                            continue
                        if retryable:
                            raise ModelTransportError(
                                "Tinker stream remained overloaded after "
                                f"{attempt + 1} attempts: {problem}"
                            ) from None
                        raise ModelTransportError(problem) from None
                    except ModelResponseError as exc:
                        problem = redact_secrets(exc, (self._api_key,))
                        report_attempt(
                            attempt=attempt + 1,
                            status="error",
                            retryable=False,
                            http_status=response.status_code,
                            response_text=decoder.raw_text(),
                            error=problem,
                            started=attempt_started,
                            input_tokens=decoder.input_tokens,
                            output_tokens=decoder.output_tokens,
                            cache_creation_input_tokens=decoder.cache_creation_input_tokens,
                            cache_read_input_tokens=decoder.cache_read_input_tokens,
                            stop_reason=decoder.stop_reason,
                            request_id=response_request_id,
                            raw_bytes_received=progress.raw_bytes_received,
                            generated_content=decoder.generated_content,
                            message_complete=decoder.stopped,
                            thinking_text=decoder.thinking_text(),
                            request_chars=request_chars,
                            message_chars=message_chars,
                        )
                        if isinstance(exc, ModelTruncationError):
                            raise ModelTruncationError(
                                problem,
                                output_tokens=decoder.output_tokens,
                            ) from None
                        raise ModelResponseError(problem) from None
            except httpx.TransportError as exc:
                if active_response is not None:
                    progress.raw_bytes_received = max(
                        progress.raw_bytes_received,
                        active_response.num_bytes_downloaded,
                    )
                last_problem = redact_secrets(exc, (self._api_key,))
                retryable = not decoder.generated_content and not decoder.stopped
                retry_delay = _retry_delay(None, attempt, payload_fingerprint)
                report_attempt(
                    attempt=attempt + 1,
                    status=("retry" if retryable and attempt < TRANSPORT_RETRIES else "error"),
                    retryable=retryable,
                    http_status=response_status,
                    response_text=decoder.raw_text() or None,
                    error=last_problem,
                    started=attempt_started,
                    input_tokens=decoder.input_tokens,
                    output_tokens=decoder.output_tokens,
                    cache_creation_input_tokens=decoder.cache_creation_input_tokens,
                    cache_read_input_tokens=decoder.cache_read_input_tokens,
                    retry_delay_seconds=(
                        retry_delay if retryable and attempt < TRANSPORT_RETRIES else None
                    ),
                    stop_reason=decoder.stop_reason,
                    request_id=response_request_id,
                    raw_bytes_received=progress.raw_bytes_received,
                    generated_content=decoder.generated_content,
                    message_complete=decoder.stopped,
                    thinking_text=decoder.thinking_text(),
                    request_chars=request_chars,
                    message_chars=message_chars,
                )
                if not retryable:
                    raise ModelTransportError(
                        "Tinker stream failed after a partial response; "
                        f"the request was not retried: {last_problem}"
                    ) from None
                if attempt >= TRANSPORT_RETRIES:
                    raise ModelTransportError(
                        f"Tinker transport failed after {attempt + 1} attempts: {last_problem}"
                    ) from None
                self._sleep(retry_delay)
                continue

            report_attempt(
                attempt=attempt + 1,
                status="success",
                retryable=False,
                http_status=response_status,
                response_text=reply["raw_text"],
                error=None,
                started=attempt_started,
                input_tokens=reply["input_tokens"],
                output_tokens=reply["output_tokens"],
                cache_creation_input_tokens=reply["cache_creation_input_tokens"],
                cache_read_input_tokens=reply["cache_read_input_tokens"],
                stop_reason=reply["stop_reason"],
                request_id=response_request_id,
                raw_bytes_received=progress.raw_bytes_received,
                generated_content=decoder.generated_content,
                message_complete=decoder.stopped,
                answer_text=reply["text"],
                thinking_text=decoder.thinking_text(),
                request_chars=request_chars,
                message_chars=message_chars,
            )
            elapsed_ms = max(0, round((time.monotonic() - started) * 1_000))
            return ModelReply(
                text=reply["text"],
                input_tokens=reply["input_tokens"],
                output_tokens=reply["output_tokens"],
                latency_ms=elapsed_ms,
            )

        raise ModelTransportError(f"Tinker request failed: {last_problem}")

    def _report_attempt(
        self,
        callback: Callable[[dict[str, object]], None] | None,
        *,
        attempt: int,
        status: str,
        retryable: bool,
        http_status: int | None,
        response_text: str | None,
        error: str | None,
        started: float,
        max_output_tokens: int,
        reasoning_effort: bool,
        request_sha256: str,
        input_tokens: int | None = None,
        output_tokens: int | None = None,
        cache_creation_input_tokens: int | None = None,
        cache_read_input_tokens: int | None = None,
        retry_delay_seconds: float | None = None,
        stop_reason: str | None = None,
        request_id: str | None = None,
        raw_bytes_received: int = 0,
        generated_content: bool = False,
        message_complete: bool = False,
        answer_text: str | None = None,
        thinking_text: str | None = None,
        request_chars: int | None = None,
        message_chars: int | None = None,
    ) -> None:
        if callback is None:
            return
        safe_response = (
            redact_secrets(response_text, (self._api_key,)) if response_text is not None else None
        )
        safe_error = redact_secrets(error, (self._api_key,)) if error is not None else None
        callback(
            {
                "attempt": attempt,
                "model": MODEL_ID,
                "temperature": TEMPERATURE,
                "max_output_tokens": max_output_tokens,
                "reasoning_effort": reasoning_effort,
                "request_sha256": request_sha256,
                "provider": "tinker",
                "transport": "anthropic_sse",
                "status": status,
                "provider_status": status,
                "retryable": retryable,
                "http_status": http_status,
                "response": safe_response,
                "error": safe_error,
                "input_tokens": input_tokens,
                "output_tokens": output_tokens,
                "cache_creation_input_tokens": cache_creation_input_tokens,
                "cache_read_input_tokens": cache_read_input_tokens,
                "retry_delay_seconds": retry_delay_seconds,
                "transport_retry_delay_ms": (
                    max(0, round(retry_delay_seconds * 1_000))
                    if retry_delay_seconds is not None
                    else None
                ),
                "stop_reason": stop_reason,
                "request_id": (
                    redact_secrets(request_id, (self._api_key,)) if request_id is not None else None
                ),
                "raw_bytes_received": raw_bytes_received,
                "generated_content": generated_content,
                "message_complete": message_complete,
                "message_chars": message_chars,
                "request_chars": request_chars,
                "response_chars": len(response_text) if response_text is not None else None,
                "answer_chars": len(answer_text) if answer_text is not None else None,
                "thinking_chars": len(thinking_text) if thinking_text is not None else None,
                "latency_ms": max(0, round((time.monotonic() - started) * 1_000)),
            }
        )
