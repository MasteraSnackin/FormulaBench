"""Inference provider boundary for the fixed hackathon model."""

from __future__ import annotations

import asyncio
import hashlib
import json
import os
import re
import threading
from collections.abc import Mapping, Sequence
from contextlib import suppress
from dataclasses import dataclass
from functools import lru_cache
from importlib.metadata import version
from typing import Protocol

from .constants import (
    DEFAULT_MAX_OUTPUT_TOKENS,
    KEY_ENV_VAR,
    MODEL_ID,
    MODEL_PROVENANCE,
    NUM_SAMPLES,
    PROVIDER_MAX_RETRIES,
    RENDERER_ID,
    STOP_TOKEN,
    STOP_TOKEN_ID,
    TEMPERATURE,
    TINKER_TELEMETRY_ENV_VAR,
    TINKER_TELEMETRY_VALUE,
    TOKENIZER_CHAT_TEMPLATE_SHA256,
    TOKENIZER_REVISION,
)

_FATAL_HTTP_STATUSES = frozenset({400, 401, 403, 404, 422})
_SAFE_CLASS_NAME = re.compile(r"[A-Za-z_][A-Za-z0-9_]{0,79}\Z")
_SAFE_TINKER_CHECKPOINT = re.compile(
    r"tinker://[A-Za-z0-9][A-Za-z0-9._-]{0,127}"
    r"(?::train:[0-9]{1,10})?/sampler_weights/"
    r"[A-Za-z0-9][A-Za-z0-9._-]{0,127}\Z"
)
_MAX_TINKER_CHECKPOINT_LENGTH = 512
_CHECKPOINT_PROVENANCE_PREFIX = "sampler_checkpoint_sha256:"


class ProviderConfigurationError(RuntimeError):
    """Raised when the provider cannot safely accept further inference calls."""

    def __init__(
        self,
        message: str,
        *,
        status_code: int | None = None,
        underlying_exception_class: str | None = None,
    ) -> None:
        super().__init__(message)
        self.status_code = (
            status_code
            if isinstance(status_code, int)
            and not isinstance(status_code, bool)
            and 100 <= status_code <= 599
            else None
        )
        self.underlying_exception_class = (
            underlying_exception_class
            if underlying_exception_class is None
            or (
                isinstance(underlying_exception_class, str)
                and _SAFE_CLASS_NAME.fullmatch(underlying_exception_class) is not None
            )
            else "HTTPError"
        )


class ProviderResponseError(RuntimeError):
    """Raised when inference succeeds but does not return a usable answer."""

    def __init__(
        self,
        message: str,
        *,
        stop_reason: str | None = None,
        parse_termination: str | None = None,
        response_rejection: str | None = None,
        input_tokens: int | None = None,
        output_tokens: int | None = None,
    ) -> None:
        super().__init__(message)
        self.stop_reason = stop_reason
        self.parse_termination = parse_termination
        self.response_rejection = response_rejection
        self.input_tokens = input_tokens
        self.output_tokens = output_tokens

    def add_termination_context(
        self,
        *,
        stop_reason: str,
        input_tokens: int,
        output_tokens: int,
    ) -> ProviderResponseError:
        self.stop_reason = stop_reason
        self.parse_termination = "stop_sequence"
        self.input_tokens = input_tokens
        self.output_tokens = output_tokens
        return self


@dataclass(frozen=True, slots=True)
class Completion:
    text: str
    input_tokens: int
    output_tokens: int
    model: str
    renderer: str
    stop_reason: str | None = None
    parse_termination: str | None = None
    response_format: str | None = None
    model_provenance: str | None = None


class CompletionProvider(Protocol):
    model: str
    model_provenance: str
    renderer_name: str

    async def complete(self, *, system_prompt: str, user_prompt: str) -> Completion: ...


ANSWER_TOOL_NAME = "submit_spreadsheet_answer"
_FINITE_RANGE_PATTERN = (
    r"^\$?[A-Za-z]{1,3}\$?[1-9][0-9]*:"
    r"\$?[A-Za-z]{1,3}\$?[1-9][0-9]*$"
)
_DATE_VALUE_PATTERN = r"^[0-9]{4}-[0-9]{2}-[0-9]{2}$"
_DATETIME_VALUE_PATTERN = (
    r"^[0-9]{4}-[0-9]{2}-[0-9]{2}T"
    r"[0-9]{2}:[0-9]{2}:[0-9]{2}(?:\.[0-9]{1,3}0{0,3})?$"
)
_VALUE_SCHEMA = {
    "oneOf": [
        {"type": "string", "maxLength": 32_767},
        {"type": "number"},
        {"type": "boolean"},
        {"type": "null"},
        {
            "type": "object",
            "additionalProperties": False,
            "properties": {
                "type": {"const": "date"},
                "value": {"type": "string", "pattern": _DATE_VALUE_PATTERN},
            },
            "required": ["type", "value"],
        },
        {
            "type": "object",
            "additionalProperties": False,
            "properties": {
                "type": {"const": "datetime"},
                "value": {"type": "string", "pattern": _DATETIME_VALUE_PATTERN},
            },
            "required": ["type", "value"],
        },
    ],
    "description": (
        "A scalar, a timezone-free typed Excel date/datetime object, or an Excel formula "
        "beginning with '='. Date-looking strings remain text. Formulas are limited to 8192 "
        "characters; other strings are limited to 32767. Datetime fractions must resolve "
        "exactly to milliseconds."
    ),
}
ANSWER_TOOL = {
    "name": ANSWER_TOOL_NAME,
    "description": (
        "Submit exact answer coverage using cells, optional compact rectangular fills, or finite "
        "preserve ranges on existing target sheets, plus only necessary exact existing merged "
        "ranges to unmerge. A fill formula is anchored at the range's top-left cell and copied "
        "with Excel-relative reference translation; a constant repeats. A preserve range counts "
        "towards coverage while leaving its existing cells unchanged."
    ),
    "input_schema": {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "cells": {
                "type": "array",
                "items": {
                    "type": "object",
                    "additionalProperties": False,
                    "properties": {
                        "sheet": {"type": "string"},
                        "cell": {"type": "string"},
                        "value": _VALUE_SCHEMA,
                    },
                    "required": ["sheet", "cell", "value"],
                },
                "maxItems": 250_000,
            },
            "fills": {
                "type": "array",
                "items": {
                    "type": "object",
                    "additionalProperties": False,
                    "properties": {
                        "sheet": {"type": "string"},
                        "range": {"type": "string", "pattern": _FINITE_RANGE_PATTERN},
                        "value": _VALUE_SCHEMA,
                    },
                    "required": ["sheet", "range", "value"],
                },
            },
            "preserve_ranges": {
                "type": "array",
                "items": {
                    "type": "object",
                    "additionalProperties": False,
                    "properties": {
                        "sheet": {"type": "string"},
                        "range": {"type": "string", "pattern": _FINITE_RANGE_PATTERN},
                    },
                    "required": ["sheet", "range"],
                },
            },
            "unmerge_ranges": {
                "type": "array",
                "items": {
                    "type": "object",
                    "additionalProperties": False,
                    "properties": {
                        "sheet": {"type": "string"},
                        "range": {"type": "string", "pattern": _FINITE_RANGE_PATTERN},
                    },
                    "required": ["sheet", "range"],
                },
            },
        },
        "required": ["cells"],
    },
}


def require_provider_key(environ: Mapping[str, str] | None = None) -> None:
    """Fail before creating a client without ever returning or logging the secret."""

    source = os.environ if environ is None else environ
    if not source.get(KEY_ENV_VAR, "").strip():
        raise ProviderConfigurationError(
            f"{KEY_ENV_VAR} is required and must be supplied through the environment"
        )


QWEN_ANSWER_TOOL = {
    "name": ANSWER_TOOL_NAME,
    "description": ANSWER_TOOL["description"],
    "parameters": ANSWER_TOOL["input_schema"],
}
_TOOL_CALL_RE = re.compile(
    r"\A\s*<tool_call>\s*<function=([A-Za-z_][A-Za-z0-9_]*)>(.*?)"
    r"</function>\s*</tool_call>\s*\Z",
    re.DOTALL,
)
_PARAMETER_RE = re.compile(
    r"\s*<parameter=([A-Za-z_][A-Za-z0-9_]*)>\s*(.*?)\s*</parameter>",
    re.DOTALL,
)
_ANSWER_PARAMETERS = frozenset(ANSWER_TOOL["input_schema"]["properties"])


def _reject_json_constant(_value: str) -> None:
    raise ValueError("non-finite JSON value")


def _reject_duplicate_keys(pairs: list[tuple[str, object]]) -> dict[str, object]:
    result: dict[str, object] = {}
    for key, value in pairs:
        if key in result:
            raise ValueError("duplicate JSON key")
        result[key] = value
    return result


def _strict_json_loads(value: str) -> object:
    try:
        return json.loads(
            value,
            object_pairs_hook=_reject_duplicate_keys,
            parse_constant=_reject_json_constant,
            parse_float=lambda raw: _finite_float(raw),
        )
    except (json.JSONDecodeError, ValueError, OverflowError) as exc:
        raise ProviderResponseError(
            "provider answer is not strict JSON",
            response_rejection="rejected_json",
        ) from exc


def _finite_float(value: str) -> float:
    parsed = float(value)
    if parsed == float("inf") or parsed == float("-inf"):
        raise ValueError("non-finite JSON value")
    return parsed


def parse_native_answer(value: str) -> tuple[str, str]:
    """Accept only an anchored Qwen tool call or a complete strict JSON object."""

    stripped = value.strip()
    if stripped.startswith("{"):
        parsed = _strict_json_loads(stripped)
        if not isinstance(parsed, Mapping):  # Defensive: the first byte already requires an object.
            raise ProviderResponseError(
                "provider JSON answer must be an object",
                response_rejection="rejected_json_type",
            )
        return stripped, "strict_json_fallback"

    if stripped.count("<tool_call>") != 1 or stripped.count("</tool_call>") != 1:
        raise ProviderResponseError(
            "provider returned an invalid number of tool calls",
            response_rejection="rejected_tool_grammar",
        )
    match = _TOOL_CALL_RE.fullmatch(stripped)
    if match is None:
        raise ProviderResponseError(
            "provider answer does not match the required tool-call grammar",
            response_rejection="rejected_tool_grammar",
        )
    if match.group(1) != ANSWER_TOOL_NAME:
        raise ProviderResponseError(
            "provider returned an unexpected tool call",
            response_rejection="rejected_tool_name",
        )

    body = match.group(2)
    position = 0
    raw_parameters: list[tuple[str, str]] = []
    seen: set[str] = set()
    while position < len(body):
        parameter = _PARAMETER_RE.match(body, position)
        if parameter is None:
            if body[position:].strip():
                raise ProviderResponseError(
                    "provider returned malformed tool parameters",
                    response_rejection="rejected_parameter_grammar",
                )
            break
        name = parameter.group(1)
        if name not in _ANSWER_PARAMETERS:
            raise ProviderResponseError(
                "provider returned an unexpected tool parameter",
                response_rejection="rejected_parameter_name",
            )
        if name in seen:
            raise ProviderResponseError(
                "provider returned a duplicate tool parameter",
                response_rejection="rejected_duplicate_parameter",
            )
        raw_json = parameter.group(2).strip()
        _strict_json_loads(raw_json)
        raw_parameters.append((name, raw_json))
        seen.add(name)
        position = parameter.end()

    if "cells" not in seen:
        raise ProviderResponseError(
            "provider answer omitted the required cells parameter",
            response_rejection="rejected_missing_parameter",
        )
    # Retain every raw parameter value: the downstream response parser performs
    # its own duplicate-key and finite-number checks over the assembled object.
    assembled = (
        "{"
        + ",".join(
            f"{json.dumps(name, ensure_ascii=False)}:{raw_json}"
            for name, raw_json in raw_parameters
        )
        + "}"
    )
    _strict_json_loads(assembled)
    return assembled, "qwen_xml_tool_call"


def _non_retryable_http_status(exc: BaseException) -> int | None:
    status = getattr(exc, "status_code", None)
    if status is None:
        response = getattr(exc, "response", None)
        status = getattr(response, "status_code", None)
    if isinstance(status, int) and not isinstance(status, bool) and status in _FATAL_HTTP_STATUSES:
        return status
    return None


_SHARED_CLIENT_LOCK = threading.Lock()
_SHARED_SERVICE_CLIENT: object | None = None
_SHARED_SAMPLING_CLIENT: object | None = None


def _environment_flag(name: str) -> bool:
    return os.environ.get(name, "").strip().lower() in {"1", "true", "yes", "on"}


@lru_cache(maxsize=1)
def _load_tokenizer() -> object:
    os.environ.setdefault("TRANSFORMERS_VERBOSITY", "error")
    from transformers import AutoTokenizer

    options: dict[str, object] = {
        "revision": TOKENIZER_REVISION,
        "trust_remote_code": False,
    }
    if _environment_flag("HF_HUB_OFFLINE") or _environment_flag("TRANSFORMERS_OFFLINE"):
        options["local_files_only"] = True
    return AutoTokenizer.from_pretrained(MODEL_ID, **options)


def _shared_sampling_client() -> object:
    """Create exactly one native service/sampling-client pair per process."""

    global _SHARED_SAMPLING_CLIENT, _SHARED_SERVICE_CLIENT
    with _SHARED_CLIENT_LOCK:
        if _SHARED_SAMPLING_CLIENT is None:
            # Pinned Tinker 0.27.1 telemetry can attach complete request headers
            # to some error events. Force it off before importing or constructing
            # any SDK client so an inherited environment value cannot expose the
            # API key or Cloudflare access credentials.
            os.environ[TINKER_TELEMETRY_ENV_VAR] = TINKER_TELEMETRY_VALUE
            import tinker
            from tinker.lib.retry_handler import RetryConfig

            service = tinker.ServiceClient(
                api_key=os.environ[KEY_ENV_VAR],
                max_retries=PROVIDER_MAX_RETRIES,
            )
            sampler = service.create_sampling_client(
                base_model=MODEL_ID,
                retry_config=RetryConfig(enable_retry_logic=False),
            )
            _SHARED_SERVICE_CLIENT = service
            _SHARED_SAMPLING_CLIENT = sampler
        return _SHARED_SAMPLING_CLIENT


def validate_sampler_checkpoint(checkpoint: str) -> str:
    """Return a bounded canonical Tinker path without reflecting invalid input."""

    if (
        not isinstance(checkpoint, str)
        or len(checkpoint) > _MAX_TINKER_CHECKPOINT_LENGTH
        or _SAFE_TINKER_CHECKPOINT.fullmatch(checkpoint) is None
    ):
        raise ProviderConfigurationError("sampler checkpoint must be a safe tinker:// path")
    return checkpoint


def sampler_checkpoint_provenance(checkpoint: str) -> str:
    """Return a hash-only trace identifier for one valid checkpoint path."""

    validated = validate_sampler_checkpoint(checkpoint)
    digest = hashlib.sha256(validated.encode("utf-8")).hexdigest()
    return f"{_CHECKPOINT_PROVENANCE_PREFIX}{digest}"


def _checkpoint_service_client() -> object:
    """Create the service boundary owned by one checkpoint provider."""

    os.environ[TINKER_TELEMETRY_ENV_VAR] = TINKER_TELEMETRY_VALUE
    import tinker

    return tinker.ServiceClient(
        api_key=os.environ[KEY_ENV_VAR],
        max_retries=PROVIDER_MAX_RETRIES,
    )


def _checkpoint_sampling_client(service: object, checkpoint: str) -> object:
    """Create a non-shared native sampler on one owned service."""

    from tinker.lib.retry_handler import RetryConfig

    return service.create_sampling_client(
        base_model=MODEL_ID,
        model_path=checkpoint,
        retry_config=RetryConfig(enable_retry_logic=False),
    )


async def _drain_tinker_futures_poller(sampler: object) -> None:
    """Cancel and await Tinker 0.27.1's sampler poller without closing shared clients."""

    # This reaches into a pinned SDK implementation detail. Future Tinker
    # versions must be inspected before the compatibility path is extended.
    if version("tinker") != "0.27.1":
        return
    poller = getattr(sampler, "_futures_poller", None)
    holder = getattr(sampler, "holder", None)
    schedule = getattr(holder, "run_coroutine_threadsafe", None)
    if poller is None or not callable(schedule) or not callable(getattr(poller, "close", None)):
        return

    async def cancel_and_drain() -> None:
        # Both the poller's task and close() belong to the holder's dedicated
        # Tinker event loop. Keep a reference because close() clears _task.
        task = getattr(poller, "_task", None)
        if task is None:
            return
        poller.close()
        with suppress(asyncio.CancelledError):
            await task

    cleanup = cancel_and_drain()
    try:
        thread_future = schedule(cleanup)
    except BaseException:
        # A closed or failed holder loop can reject scheduling synchronously.
        # Close the unstarted coroutine so the error path cannot leak it.
        cleanup.close()
        raise
    result = getattr(thread_future, "result", None)
    if callable(result):
        # AwaitableConcurrentFuture.result() blocks, so keep it off the CLI's
        # asyncio loop while the Tinker loop drains the cancelled task.
        await asyncio.to_thread(result)
    else:
        await thread_future


async def _drain_tinker_service_client(service: object) -> None:
    """Await pinned SDK background-task cleanup for a run-local service holder."""

    if version("tinker") != "0.27.1":
        return
    holder = getattr(service, "_session_holder", None)
    cleanup_method = getattr(holder, "_async_cleanup", None)
    schedule = getattr(holder, "run_coroutine_threadsafe", None)
    if holder is None or not callable(cleanup_method) or not callable(schedule):
        return

    cleanup = cleanup_method()
    try:
        thread_future = schedule(cleanup)
    except BaseException:
        cleanup.close()
        raise
    result = getattr(thread_future, "result", None)
    if callable(result):
        await asyncio.to_thread(result)
    else:
        await thread_future


def _render_prompt(tokenizer: object, system_prompt: str, user_prompt: str) -> list[int]:
    apply_chat_template = getattr(tokenizer, "apply_chat_template", None)
    if apply_chat_template is None:
        raise ProviderConfigurationError("pinned tokenizer has no chat template")
    tokens = apply_chat_template(
        [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
        tools=[QWEN_ANSWER_TOOL],
        tokenize=True,
        add_generation_prompt=True,
        enable_thinking=False,
        return_dict=False,
    )
    if (
        not isinstance(tokens, list)
        or not tokens
        or any(isinstance(token, bool) or not isinstance(token, int) for token in tokens)
    ):
        raise ProviderConfigurationError("pinned tokenizer returned invalid prompt tokens")
    return tokens


def validate_native_runtime() -> None:
    """Validate pinned local dependencies and rendering without a provider call."""

    if version("tinker") != "0.27.1":
        raise ProviderConfigurationError("unexpected Tinker SDK version")
    if version("transformers") != "5.5.4" or version("tokenizers") != "0.22.2":
        raise ProviderConfigurationError("unexpected tokenizer runtime version")
    tokenizer = _load_tokenizer()
    chat_template = getattr(tokenizer, "chat_template", None)
    if (
        not isinstance(chat_template, str)
        or hashlib.sha256(chat_template.encode("utf-8")).hexdigest()
        != TOKENIZER_CHAT_TEMPLATE_SHA256
    ):
        raise ProviderConfigurationError("pinned tokenizer chat template failed verification")
    token_id = getattr(tokenizer, "convert_tokens_to_ids", lambda _token: None)(STOP_TOKEN)
    if token_id != STOP_TOKEN_ID:
        raise ProviderConfigurationError("pinned tokenizer stop token failed verification")
    _render_prompt(tokenizer, "preflight system", "preflight user")


class TinkerProvider:
    """One-call provider using Tinker's native sampling client."""

    model = MODEL_ID

    def __init__(
        self,
        *,
        max_output_tokens: int = DEFAULT_MAX_OUTPUT_TOKENS,
        sampler_checkpoint: str | None = None,
        sampling_client: object | None = None,
        tokenizer: object | None = None,
    ) -> None:
        os.environ[TINKER_TELEMETRY_ENV_VAR] = TINKER_TELEMETRY_VALUE
        require_provider_key()
        if max_output_tokens < 1:
            raise ValueError("max_output_tokens must be positive")
        self.renderer_name = RENDERER_ID
        self.model_provenance = (
            MODEL_PROVENANCE
            if sampler_checkpoint is None
            else sampler_checkpoint_provenance(sampler_checkpoint)
        )
        self._sampler_checkpoint = sampler_checkpoint
        self._sampling_client = sampling_client
        self._service_client: object | None = None
        self._uses_shared_sampling_client = sampling_client is None and sampler_checkpoint is None
        self._drains_native_sampling_client = sampling_client is None
        self._client_creation_lock = asyncio.Lock()
        self._tokenizer = tokenizer if tokenizer is not None else _load_tokenizer()
        self._max_output_tokens = max_output_tokens
        self._fatal_lock = threading.Lock()
        self._fatal_failure: tuple[int, str] | None = None

    def _raise_if_fatal(self) -> None:
        with self._fatal_lock:
            failure = self._fatal_failure
        if failure is not None:
            status, class_name = failure
            raise ProviderConfigurationError(
                f"provider circuit is open after non-retryable HTTP {status} ({class_name})",
                status_code=status,
                underlying_exception_class=class_name,
            )

    async def complete(self, *, system_prompt: str, user_prompt: str) -> Completion:
        self._raise_if_fatal()
        prompt_tokens = _render_prompt(self._tokenizer, system_prompt, user_prompt)
        try:
            import tinker

            sampler = self._sampling_client
            if sampler is None:
                async with self._client_creation_lock:
                    sampler = self._sampling_client
                    if sampler is None:
                        if self._sampler_checkpoint is None:
                            sampler = await asyncio.to_thread(_shared_sampling_client)
                        else:
                            service = await asyncio.to_thread(_checkpoint_service_client)
                            self._service_client = service
                            try:
                                sampler = await asyncio.to_thread(
                                    _checkpoint_sampling_client,
                                    service,
                                    self._sampler_checkpoint,
                                )
                            except BaseException:
                                try:
                                    await _drain_tinker_service_client(service)
                                finally:
                                    self._service_client = None
                                raise
                        self._sampling_client = sampler
            response = await sampler.sample_async(
                prompt=tinker.ModelInput.from_ints(prompt_tokens),
                num_samples=NUM_SAMPLES,
                sampling_params=tinker.SamplingParams(
                    max_tokens=self._max_output_tokens,
                    temperature=TEMPERATURE,
                    stop=[STOP_TOKEN_ID],
                ),
            )
        except Exception as exc:
            if (status := _non_retryable_http_status(exc)) is not None:
                class_name = type(exc).__name__
                if _SAFE_CLASS_NAME.fullmatch(class_name) is None:
                    class_name = "HTTPError"
                with self._fatal_lock:
                    if self._fatal_failure is None:
                        self._fatal_failure = (status, class_name)
                    stored_status, stored_class = self._fatal_failure
                raise ProviderConfigurationError(
                    "provider circuit is open after non-retryable HTTP "
                    f"{stored_status} ({stored_class})",
                    status_code=stored_status,
                    underlying_exception_class=stored_class,
                ) from None
            raise
        sequences = getattr(response, "sequences", None)
        if not isinstance(sequences, Sequence) or isinstance(sequences, (str, bytes)):
            raise ProviderResponseError(
                "provider returned invalid sequence metadata",
                response_rejection="rejected_sequence_metadata",
                input_tokens=len(prompt_tokens),
            )
        if len(sequences) != NUM_SAMPLES:
            raise ProviderResponseError(
                "provider returned an unexpected number of sequences",
                response_rejection="rejected_sequence_count",
                input_tokens=len(prompt_tokens),
            )
        sequence = sequences[0]
        stop_reason = getattr(sequence, "stop_reason", None)
        safe_stop_reason = stop_reason if stop_reason in {"stop", "length"} else None
        tokens = getattr(sequence, "tokens", None)
        if not isinstance(tokens, list) or any(
            isinstance(token, bool) or not isinstance(token, int) for token in tokens
        ):
            raise ProviderResponseError(
                "provider returned invalid output tokens",
                stop_reason=safe_stop_reason,
                response_rejection="rejected_output_tokens",
                input_tokens=len(prompt_tokens),
            )
        output_tokens = len(tokens)
        if stop_reason != "stop":
            raise ProviderResponseError(
                "provider sequence did not stop at the configured token",
                stop_reason=safe_stop_reason,
                response_rejection="rejected_stop_reason",
                input_tokens=len(prompt_tokens),
                output_tokens=output_tokens,
            )
        if not tokens or tokens[-1] != STOP_TOKEN_ID or tokens.count(STOP_TOKEN_ID) != 1:
            raise ProviderResponseError(
                "provider sequence has invalid stop-token termination",
                stop_reason="stop",
                response_rejection="rejected_stop_token",
                input_tokens=len(prompt_tokens),
                output_tokens=output_tokens,
            )
        decoded = self._tokenizer.decode(
            tokens[:-1],
            skip_special_tokens=False,
            clean_up_tokenization_spaces=False,
        )
        if not isinstance(decoded, str):
            raise ProviderResponseError(
                "pinned tokenizer returned invalid decoded output",
                stop_reason="stop",
                parse_termination="stop_sequence",
                response_rejection="rejected_decoded_output",
                input_tokens=len(prompt_tokens),
                output_tokens=output_tokens,
            )
        try:
            answer_text, response_format = parse_native_answer(decoded)
        except ProviderResponseError as exc:
            raise exc.add_termination_context(
                stop_reason="stop",
                input_tokens=len(prompt_tokens),
                output_tokens=output_tokens,
            ) from None
        return Completion(
            text=answer_text,
            input_tokens=len(prompt_tokens),
            output_tokens=output_tokens,
            model=MODEL_ID,
            renderer=self.renderer_name,
            stop_reason="stop",
            parse_termination="stop_sequence",
            response_format=response_format,
            model_provenance=self.model_provenance,
        )

    async def close(self) -> None:
        # Externally injected clients remain caller-owned. The process-scoped
        # base sampler is retained for reuse; checkpoint services are run-local.
        if not self._drains_native_sampling_client or self._sampling_client is None:
            return
        try:
            await _drain_tinker_futures_poller(self._sampling_client)
        finally:
            if not self._uses_shared_sampling_client:
                service = self._service_client
                try:
                    if service is not None:
                        await _drain_tinker_service_client(service)
                finally:
                    self._sampling_client = None
                    self._service_client = None
