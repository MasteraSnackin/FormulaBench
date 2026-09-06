"""Bounded, dry-run-first LoRA supervised fine-tuning for FormulaBench.

The runner accepts only the versioned corpus emitted by ``training.corpus``.
Dry-run validation and token accounting happen before any Tinker client exists.
Paid training requires the explicit ``--execute`` flag and a project ID.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import math
import os
import random
import re
import sys
from collections.abc import Callable, Mapping, Sequence
from dataclasses import dataclass
from datetime import UTC, datetime
from importlib.metadata import PackageNotFoundError, version
from pathlib import Path
from typing import Any

from openpyxl.utils.cell import get_column_letter, range_boundaries

from formulabench.constants import (
    KEY_ENV_VAR,
    MODEL_ID,
    TINKER_TELEMETRY_ENV_VAR,
    TINKER_TELEMETRY_VALUE,
    TOKENIZER_REVISION,
)
from formulabench.contract import SpreadsheetResponse
from formulabench.prompts import SYSTEM_PROMPT
from formulabench.provider import ANSWER_TOOL_NAME, QWEN_ANSWER_TOOL

ROOT = Path(__file__).resolve().parents[1]
DEFAULT_CORPUS_MANIFEST = ROOT / "training" / "generated" / "corpus_manifest.json"
DEFAULT_SPLIT_MANIFEST = ROOT / "experiments" / "public_split.json"

CORPUS_SCHEMA_VERSION = 1
CORPUS_BUILDER = "FormulaBench/sft-corpus/v1"
PARTITION_SEED = "FormulaBench/sft-partition/v1"
RENDERER_NAME = "qwen3_8_disable_thinking"
LOSS_FUNCTION = "cross_entropy"
TRAIN_ON = "last_assistant_message"
MAX_DEVELOPMENT_EXAMPLES = 80
MAX_TARGET_CELLS = 500
MAX_CORPUS_BYTES = 128 * 1024 * 1024
MAX_MANIFEST_BYTES = 4 * 1024 * 1024
MAX_STEPS = 200
MAX_EPOCHS = 10
MAX_BATCH_SIZE = 16
_SHA256_RE = re.compile(r"^[0-9a-f]{64}$")
_CHECKPOINT_NAME_RE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9._-]{0,63}$")
_ROW_KEYS = {
    "answer_cell_count",
    "messages",
    "partition",
    "provenance",
    "schema_version",
    "split",
    "target_cell_count",
    "task_id",
}


class TrainingInputError(ValueError):
    """A local corpus, configuration, or tokenisation preflight failed."""


class TrainingExecutionError(RuntimeError):
    """Paid training failed; provider details are deliberately not echoed."""


@dataclass(frozen=True, slots=True)
class CorpusRow:
    task_id: str
    partition: str
    messages: tuple[dict[str, str], ...]
    target_cell_count: int
    answer_cell_count: int
    raw: dict[str, Any]


@dataclass(frozen=True, slots=True)
class ValidatedCorpus:
    manifest_path: Path
    corpus_path: Path
    corpus_sha256: str
    split_manifest_sha256: str
    dataset_manifest_sha256: str
    development_golden_binding_sha256: str
    rows: tuple[CorpusRow, ...]
    excluded_development_count: int

    @property
    def train_rows(self) -> tuple[CorpusRow, ...]:
        return tuple(row for row in self.rows if row.partition == "train")

    @property
    def validation_rows(self) -> tuple[CorpusRow, ...]:
        return tuple(row for row in self.rows if row.partition == "validation")


@dataclass(frozen=True, slots=True)
class PreparedExample:
    task_id: str
    partition: str
    datum: Any
    total_tokens: int
    completion_tokens: int


@dataclass(frozen=True, slots=True)
class PreparedCorpus:
    corpus: ValidatedCorpus
    examples: tuple[PreparedExample, ...]
    wire_format_match: bool

    @property
    def train_examples(self) -> tuple[PreparedExample, ...]:
        return tuple(item for item in self.examples if item.partition == "train")

    @property
    def validation_examples(self) -> tuple[PreparedExample, ...]:
        return tuple(item for item in self.examples if item.partition == "validation")


@dataclass(frozen=True, slots=True)
class TrainConfig:
    manifest_path: Path
    split_manifest_path: Path
    execute: bool
    project_id: str | None
    run_dir: Path | None
    rank: int
    learning_rate: float
    batch_size: int
    epochs: int
    max_steps: int
    max_sequence_tokens: int
    seed: int
    beta1: float
    beta2: float
    eps: float
    weight_decay: float
    grad_clip_norm: float
    state_ttl_seconds: int
    checkpoint_prefix: str
    allow_tokenizer_download: bool


def _canonical_json(value: object) -> str:
    return json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":"))


def _sha256_bytes(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def _sha256_text(value: str) -> str:
    return _sha256_bytes(value.encode("utf-8"))


def _require_mapping(value: object, label: str) -> Mapping[str, Any]:
    if not isinstance(value, Mapping):
        raise TrainingInputError(f"{label} must be a JSON object")
    return value


def _require_text(value: object, label: str) -> str:
    if not isinstance(value, str) or not value:
        raise TrainingInputError(f"{label} must be a non-empty string")
    return value


def _require_int(value: object, label: str, *, minimum: int = 0) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value < minimum:
        raise TrainingInputError(f"{label} must be an integer >= {minimum}")
    return value


def _require_sha256(value: object, label: str) -> str:
    text = _require_text(value, label)
    if _SHA256_RE.fullmatch(text) is None:
        raise TrainingInputError(f"{label} must be a lowercase SHA-256 digest")
    return text


def _read_bounded(path: Path, maximum: int, label: str) -> bytes:
    if path.is_symlink():
        raise TrainingInputError(f"{label} must be a regular non-symlink file")
    try:
        resolved = path.resolve(strict=True)
    except (OSError, RuntimeError) as exc:
        raise TrainingInputError(f"{label} is not a readable file") from exc
    if not resolved.is_file():
        raise TrainingInputError(f"{label} must be a regular non-symlink file")
    size = resolved.stat().st_size
    if size <= 0 or size > maximum:
        raise TrainingInputError(f"{label} size must be between 1 and {maximum} bytes")
    return resolved.read_bytes()


def _load_json_file(path: Path, maximum: int, label: str) -> tuple[Mapping[str, Any], bytes]:
    raw = _read_bounded(path, maximum, label)
    try:
        payload = json.loads(raw.decode("utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise TrainingInputError(f"{label} must contain valid UTF-8 JSON") from exc
    return _require_mapping(payload, label), raw


def _resolve_corpus_path(manifest_path: Path, corpus_file: object) -> Path:
    relative = Path(_require_text(corpus_file, "manifest.corpus_file"))
    if relative.is_absolute() or ".." in relative.parts:
        raise TrainingInputError("manifest.corpus_file must be a confined relative path")
    root = manifest_path.resolve().parent
    candidate = (root / relative).resolve(strict=False)
    if candidate.parent != root:
        raise TrainingInputError("manifest.corpus_file must resolve beside the manifest")
    return candidate


def _answer_coverage_count(answer: SpreadsheetResponse, task_id: str) -> int:
    covered: set[tuple[str, str]] = set()

    def add(sheet: str, cell: str) -> None:
        key = (sheet, cell)
        if key in covered:
            raise TrainingInputError(f"corpus row {task_id} has overlapping answer coverage")
        covered.add(key)

    for cell in answer.cells:
        add(cell.sheet, cell.cell)
    for item in (*answer.fills, *answer.preserve_ranges):
        min_col, min_row, max_col, max_row = range_boundaries(item.range)
        for row in range(min_row, max_row + 1):
            for column in range(min_col, max_col + 1):
                add(item.sheet, f"{get_column_letter(column)}{row}")
    return len(covered)


def _parse_messages(row: Mapping[str, Any], task_id: str) -> tuple[dict[str, str], ...]:
    messages = row.get("messages")
    if not isinstance(messages, list) or len(messages) != 3:
        raise TrainingInputError(f"corpus row {task_id} must contain exactly three messages")
    expected_roles = ("system", "user", "assistant")
    parsed: list[dict[str, str]] = []
    for index, (message, expected_role) in enumerate(zip(messages, expected_roles, strict=True)):
        item = _require_mapping(message, f"corpus row {task_id} message {index}")
        if set(item) != {"role", "content"}:
            raise TrainingInputError(
                f"corpus row {task_id} message {index} must contain only role and content"
            )
        role = _require_text(item.get("role"), f"corpus row {task_id} message {index}.role")
        content = _require_text(
            item.get("content"), f"corpus row {task_id} message {index}.content"
        )
        if role != expected_role:
            raise TrainingInputError(
                f"corpus row {task_id} message {index} must have role {expected_role}"
            )
        parsed.append({"role": role, "content": content})
    if parsed[0]["content"] != SYSTEM_PROMPT:
        raise TrainingInputError(f"corpus row {task_id} does not use the production system prompt")
    return tuple(parsed)


def _parse_corpus_rows(raw: bytes) -> tuple[tuple[CorpusRow, ...], tuple[str, ...]]:
    if b"\r" in raw or not raw.endswith(b"\n"):
        raise TrainingInputError("corpus must be canonical LF-terminated JSONL")
    rows: list[CorpusRow] = []
    row_hashes: list[str] = []
    for line_number, line in enumerate(raw.splitlines(keepends=True), start=1):
        if not line.endswith(b"\n") or line == b"\n":
            raise TrainingInputError(f"corpus line {line_number} is not canonical JSONL")
        encoded = line[:-1]
        try:
            payload = json.loads(encoded.decode("utf-8"))
        except (UnicodeDecodeError, json.JSONDecodeError) as exc:
            raise TrainingInputError(f"corpus line {line_number} is invalid UTF-8 JSON") from exc
        row = _require_mapping(payload, f"corpus line {line_number}")
        if encoded != _canonical_json(row).encode("utf-8"):
            raise TrainingInputError(f"corpus line {line_number} is not canonical JSON")
        if set(row) != _ROW_KEYS:
            raise TrainingInputError(f"corpus line {line_number} has unexpected fields")
        if row.get("schema_version") != CORPUS_SCHEMA_VERSION:
            raise TrainingInputError(f"corpus line {line_number} has unsupported schema_version")
        task_id = _require_text(row.get("task_id"), f"corpus line {line_number}.task_id")
        if row.get("split") != "development":
            raise TrainingInputError(f"corpus row {task_id} is not development-only")
        partition = row.get("partition")
        if partition not in {"train", "validation"}:
            raise TrainingInputError(f"corpus row {task_id} has an invalid partition")
        target_count = _require_int(
            row.get("target_cell_count"), f"corpus row {task_id}.target_cell_count", minimum=1
        )
        answer_count = _require_int(
            row.get("answer_cell_count"), f"corpus row {task_id}.answer_cell_count", minimum=1
        )
        if target_count != answer_count or target_count > MAX_TARGET_CELLS:
            raise TrainingInputError(
                f"corpus row {task_id} is incomplete or exceeds the cell limit"
            )
        messages = _parse_messages(row, task_id)
        answer_text = messages[-1]["content"]
        try:
            answer_payload = json.loads(answer_text)
            answer = SpreadsheetResponse.model_validate(answer_payload, strict=True)
        except (json.JSONDecodeError, ValueError) as exc:
            raise TrainingInputError(
                f"corpus row {task_id} has an invalid assistant answer"
            ) from exc
        if answer_text != _canonical_json(answer_payload):
            raise TrainingInputError(f"corpus row {task_id} assistant answer is not canonical JSON")
        if _answer_coverage_count(answer, task_id) != answer_count:
            raise TrainingInputError(f"corpus row {task_id} answer_cell_count is not reproducible")
        provenance = _require_mapping(row.get("provenance"), f"corpus row {task_id}.provenance")
        for key in (
            "dataset_manifest_sha256",
            "split_manifest_sha256",
            "initial_workbook_sha256",
            "golden_workbook_sha256",
            "prompt_sha256",
            "answer_sha256",
        ):
            _require_sha256(provenance.get(key), f"corpus row {task_id}.provenance.{key}")
        if provenance["prompt_sha256"] != _sha256_text(messages[1]["content"]):
            raise TrainingInputError(f"corpus row {task_id} prompt digest does not match")
        if provenance["answer_sha256"] != _sha256_text(answer_text):
            raise TrainingInputError(f"corpus row {task_id} answer digest does not match")
        rows.append(
            CorpusRow(
                task_id=task_id,
                partition=partition,
                messages=messages,
                target_cell_count=target_count,
                answer_cell_count=answer_count,
                raw=dict(row),
            )
        )
        row_hashes.append(_sha256_bytes(encoded))
    if not rows or len(rows) > MAX_DEVELOPMENT_EXAMPLES:
        raise TrainingInputError("corpus must contain between 1 and 80 examples")
    return tuple(rows), tuple(row_hashes)


def load_validated_corpus(manifest_path: Path, split_manifest_path: Path) -> ValidatedCorpus:
    """Load and revalidate every corpus binding without opening a workbook."""

    manifest, _ = _load_json_file(manifest_path, MAX_MANIFEST_BYTES, "corpus manifest")
    split_manifest, split_raw = _load_json_file(
        split_manifest_path, MAX_MANIFEST_BYTES, "split manifest"
    )
    if manifest.get("schema_version") != CORPUS_SCHEMA_VERSION:
        raise TrainingInputError("corpus manifest has unsupported schema_version")
    if manifest.get("builder") != CORPUS_BUILDER:
        raise TrainingInputError("corpus manifest was not produced by the approved builder")
    if manifest.get("split") != "development" or manifest.get("verified") is not True:
        raise TrainingInputError("corpus manifest is not verified development-only data")
    if manifest.get("partition_seed") != PARTITION_SEED:
        raise TrainingInputError("corpus manifest has an unexpected partition seed")
    _require_text(manifest.get("partition_method"), "manifest.partition_method")
    if manifest.get("max_target_cells") != MAX_TARGET_CELLS:
        raise TrainingInputError("corpus manifest has an unexpected target-cell limit")
    if manifest.get("system_prompt_sha256") != _sha256_text(SYSTEM_PROMPT):
        raise TrainingInputError("corpus manifest does not bind the production system prompt")

    split_digest = _sha256_bytes(split_raw)
    if (
        _require_sha256(manifest.get("split_manifest_sha256"), "manifest.split_manifest_sha256")
        != split_digest
    ):
        raise TrainingInputError("corpus manifest does not bind the selected split manifest")
    dataset_digest = _require_sha256(
        manifest.get("dataset_manifest_sha256"), "manifest.dataset_manifest_sha256"
    )
    if split_manifest.get("dataset_manifest_sha256") != dataset_digest:
        raise TrainingInputError("dataset manifest digest disagrees with the frozen split")
    initial_binding = _require_mapping(
        split_manifest.get("initial_workbook_binding"), "split.initial_workbook_binding"
    )
    if manifest.get("initial_workbook_binding_sha256") != initial_binding.get("aggregate_sha256"):
        raise TrainingInputError("initial-workbook binding disagrees with the frozen split")
    golden_binding = _require_sha256(
        manifest.get("development_golden_binding_sha256"),
        "manifest.development_golden_binding_sha256",
    )

    development_ids = split_manifest.get("development_ids")
    if (
        not isinstance(development_ids, list)
        or len(development_ids) != MAX_DEVELOPMENT_EXAMPLES
        or any(not isinstance(item, str) or not item for item in development_ids)
        or len(set(development_ids)) != len(development_ids)
    ):
        raise TrainingInputError("split manifest must contain 80 unique development IDs")
    corpus_path = _resolve_corpus_path(manifest_path, manifest.get("corpus_file"))
    corpus_raw = _read_bounded(corpus_path, MAX_CORPUS_BYTES, "corpus")
    corpus_digest = _sha256_bytes(corpus_raw)
    if _require_sha256(manifest.get("corpus_sha256"), "manifest.corpus_sha256") != corpus_digest:
        raise TrainingInputError("corpus digest does not match the manifest")
    rows, row_hashes = _parse_corpus_rows(corpus_raw)

    ordered_ids = manifest.get("ordered_task_ids")
    if not isinstance(ordered_ids, list) or ordered_ids != [row.task_id for row in rows]:
        raise TrainingInputError("manifest ordered_task_ids do not match corpus order")
    if len(set(ordered_ids)) != len(ordered_ids):
        raise TrainingInputError("corpus contains duplicate task IDs")
    allowed = set(development_ids)
    if any(task_id not in allowed for task_id in ordered_ids):
        raise TrainingInputError("corpus contains a non-development task")
    if [task_id for task_id in development_ids if task_id in set(ordered_ids)] != ordered_ids:
        raise TrainingInputError("corpus rows do not preserve frozen development order")
    if manifest.get("example_count") != len(rows):
        raise TrainingInputError("manifest example_count does not match the corpus")

    expected_examples = manifest.get("examples")
    if not isinstance(expected_examples, list) or len(expected_examples) != len(rows):
        raise TrainingInputError("manifest examples do not match the corpus")
    for row, row_hash, expected in zip(rows, row_hashes, expected_examples, strict=True):
        item = _require_mapping(expected, f"manifest example {row.task_id}")
        provenance = _require_mapping(row.raw["provenance"], f"corpus row {row.task_id}.provenance")
        comparisons = {
            "task_id": row.task_id,
            "partition": row.partition,
            "row_sha256": row_hash,
            "initial_workbook_sha256": provenance["initial_workbook_sha256"],
            "golden_workbook_sha256": provenance["golden_workbook_sha256"],
            "prompt_sha256": provenance["prompt_sha256"],
            "answer_sha256": provenance["answer_sha256"],
            "target_cell_count": row.target_cell_count,
            "answer_cell_count": row.answer_cell_count,
        }
        if any(item.get(key) != value for key, value in comparisons.items()):
            raise TrainingInputError(
                f"manifest example {row.task_id} does not match its corpus row"
            )
        if provenance["dataset_manifest_sha256"] != dataset_digest:
            raise TrainingInputError(f"corpus row {row.task_id} has the wrong dataset binding")
        if provenance["split_manifest_sha256"] != split_digest:
            raise TrainingInputError(f"corpus row {row.task_id} has the wrong split binding")

    partition_counts = _require_mapping(
        manifest.get("partition_counts"), "manifest.partition_counts"
    )
    actual_counts = {
        "train": sum(row.partition == "train" for row in rows),
        "validation": sum(row.partition == "validation" for row in rows),
    }
    if dict(partition_counts) != actual_counts or not all(actual_counts.values()):
        raise TrainingInputError("manifest partition counts are invalid")
    eligible = _require_int(
        manifest.get("eligible_development_count"), "manifest.eligible_development_count", minimum=1
    )
    excluded = _require_int(
        manifest.get("excluded_development_count"), "manifest.excluded_development_count"
    )
    if eligible != len(rows) or eligible + excluded != len(development_ids):
        raise TrainingInputError("eligible/excluded development counts are inconsistent")
    exclusions = manifest.get("exclusions")
    if not isinstance(exclusions, list) or len(exclusions) != excluded:
        raise TrainingInputError("manifest exclusions do not match excluded count")
    excluded_ids: list[str] = []
    for index, exclusion in enumerate(exclusions):
        item = _require_mapping(exclusion, f"manifest exclusion {index}")
        task_id = _require_text(item.get("task_id"), f"manifest exclusion {index}.task_id")
        count = _require_int(
            item.get("target_cell_count"),
            f"manifest exclusion {task_id}.target_cell_count",
            minimum=1,
        )
        if item.get("reason") != "target_cell_limit_exceeded" or count <= MAX_TARGET_CELLS:
            raise TrainingInputError(f"manifest exclusion {task_id} is invalid")
        excluded_ids.append(task_id)
    if len(set(excluded_ids)) != len(excluded_ids):
        raise TrainingInputError("manifest exclusions contain duplicate task IDs")
    if set(excluded_ids) != allowed - set(ordered_ids):
        raise TrainingInputError("manifest exclusions do not cover omitted development tasks")

    return ValidatedCorpus(
        manifest_path=manifest_path.resolve(),
        corpus_path=corpus_path.resolve(),
        corpus_sha256=corpus_digest,
        split_manifest_sha256=split_digest,
        dataset_manifest_sha256=dataset_digest,
        development_golden_binding_sha256=golden_binding,
        rows=rows,
        excluded_development_count=excluded,
    )


def _default_tokenisation_dependencies(
    allow_download: bool,
) -> tuple[Any, Any, Callable[..., Any], Callable[[str], Any]]:
    try:
        from tinker_cookbook.renderers import TrainOnWhat, get_renderer
        from tinker_cookbook.renderers.base import ToolCall
        from tinker_cookbook.supervised.data import conversation_to_datum
        from transformers import AutoTokenizer
    except ImportError as exc:
        raise TrainingInputError(
            "native training dependencies are missing; install the native-tinker extra"
        ) from exc
    try:
        tokenizer = AutoTokenizer.from_pretrained(
            MODEL_ID,
            revision=TOKENIZER_REVISION,
            local_files_only=not allow_download,
        )
    except Exception as exc:
        raise TrainingInputError(
            "the pinned Qwen3.8 tokenizer is unavailable locally; prefetch it or pass "
            "--allow-tokenizer-download"
        ) from exc
    renderer = get_renderer(RENDERER_NAME, tokenizer)

    def tool_call(answer: str) -> Any:
        return ToolCall(
            function=ToolCall.FunctionBody(name=ANSWER_TOOL_NAME, arguments=answer),
            id=None,
        )

    def convert(messages: list[dict[str, Any]], selected_renderer: Any) -> Any:
        return conversation_to_datum(
            messages,
            selected_renderer,
            max_length=None,
            train_on_what=TrainOnWhat.LAST_ASSISTANT_MESSAGE,
        )

    return tokenizer, renderer, convert, tool_call


def _weights(datum: Any, task_id: str) -> list[float]:
    try:
        raw = datum.loss_fn_inputs["weights"]
        values = raw.data if hasattr(raw, "data") else raw
        result = [float(value) for value in values]
    except (AttributeError, KeyError, TypeError, ValueError) as exc:
        raise TrainingInputError(f"tokenised row {task_id} has no valid loss mask") from exc
    if any(not math.isfinite(value) or value < 0 for value in result):
        raise TrainingInputError(f"tokenised row {task_id} has an invalid loss mask")
    return result


def prepare_corpus(
    corpus: ValidatedCorpus,
    *,
    max_sequence_tokens: int,
    allow_tokenizer_download: bool = False,
    dependencies: tuple[Any, Any, Callable[..., Any], Callable[[str], Any]] | None = None,
) -> PreparedCorpus:
    """Render tool-call SFT examples with no truncation and account for both partitions."""

    tokenizer, renderer, convert, make_tool_call = (
        dependencies or _default_tokenisation_dependencies(allow_tokenizer_download)
    )
    prepared: list[PreparedExample] = []
    wire_format_match = True
    for row in corpus.rows:
        system_message, user_message, assistant_message = row.messages
        prefix = renderer.create_conversation_prefix_with_tools(
            [QWEN_ANSWER_TOOL], system_prompt=system_message["content"]
        )
        prompt_messages = [*prefix, {"role": "user", "content": user_message["content"]}]
        generation_prompt = renderer.build_generation_prompt(prompt_messages)
        production_tokens = tokenizer.apply_chat_template(
            [system_message, user_message],
            tools=[QWEN_ANSWER_TOOL],
            tokenize=True,
            add_generation_prompt=True,
            enable_thinking=False,
            return_dict=False,
        )
        cookbook_tokens = generation_prompt.to_ints()
        wire_format_match = production_tokens == cookbook_tokens
        if not wire_format_match:
            raise TrainingInputError(
                f"cookbook tool-call prompt does not match production wire format for {row.task_id}"
            )
        messages = [
            *prompt_messages,
            {
                "role": "assistant",
                "content": "",
                "tool_calls": [make_tool_call(assistant_message["content"])],
            },
        ]
        datum = convert(messages, renderer)
        try:
            total_tokens = int(datum.model_input.length)
        except (AttributeError, TypeError, ValueError) as exc:
            raise TrainingInputError(f"tokenised row {row.task_id} has no valid length") from exc
        if total_tokens <= 0 or total_tokens > max_sequence_tokens:
            raise TrainingInputError(
                f"tokenised row {row.task_id} has {total_tokens} tokens; limit is "
                f"{max_sequence_tokens}; refusing truncation"
            )
        weights = _weights(datum, row.task_id)
        if len(weights) != total_tokens:
            raise TrainingInputError(f"tokenised row {row.task_id} has a misaligned loss mask")
        positive = [position for position, weight in enumerate(weights) if weight > 0]
        prompt_tokens = int(generation_prompt.length)
        if not positive or positive[0] < prompt_tokens - 1:
            raise TrainingInputError(
                f"tokenised row {row.task_id} trains prompt tokens or has no assistant tokens"
            )
        prepared.append(
            PreparedExample(
                task_id=row.task_id,
                partition=row.partition,
                datum=datum,
                total_tokens=total_tokens,
                completion_tokens=len(positive),
            )
        )
    return PreparedCorpus(
        corpus=corpus, examples=tuple(prepared), wire_format_match=wire_format_match
    )


def _token_summary(examples: Sequence[PreparedExample]) -> dict[str, int]:
    return {
        "examples": len(examples),
        "total_tokens": sum(item.total_tokens for item in examples),
        "assistant_completion_tokens": sum(item.completion_tokens for item in examples),
        "max_sequence_tokens": max((item.total_tokens for item in examples), default=0),
    }


def _planned_batches(
    config: TrainConfig, examples: Sequence[PreparedExample]
) -> list[list[PreparedExample]]:
    batches: list[list[PreparedExample]] = []
    for epoch in range(config.epochs):
        shuffled = list(examples)
        random.Random(config.seed + epoch).shuffle(shuffled)
        for start in range(0, len(shuffled), config.batch_size):
            batches.append(shuffled[start : start + config.batch_size])
            if len(batches) >= config.max_steps:
                return batches
    return batches


def build_plan(config: TrainConfig, prepared: PreparedCorpus) -> dict[str, Any]:
    batches = _planned_batches(config, prepared.train_examples)
    return {
        "mode": "execute" if config.execute else "dry-run",
        "provider_calls": "enabled" if config.execute else "zero",
        "base_model": MODEL_ID,
        "renderer": RENDERER_NAME,
        "tokenizer_revision": TOKENIZER_REVISION,
        "tool_name": ANSWER_TOOL_NAME,
        "wire_format_matches_production": prepared.wire_format_match,
        "loss_function": LOSS_FUNCTION,
        "train_on": TRAIN_ON,
        "corpus_sha256": prepared.corpus.corpus_sha256,
        "split_manifest_sha256": prepared.corpus.split_manifest_sha256,
        "dataset_manifest_sha256": prepared.corpus.dataset_manifest_sha256,
        "development_golden_binding_sha256": (prepared.corpus.development_golden_binding_sha256),
        "train": _token_summary(prepared.train_examples),
        "validation": _token_summary(prepared.validation_examples),
        "excluded_development_examples": prepared.corpus.excluded_development_count,
        "training": {
            "rank": config.rank,
            "learning_rate": config.learning_rate,
            "batch_size": config.batch_size,
            "epochs": config.epochs,
            "max_steps": config.max_steps,
            "planned_steps": len(batches),
            "max_sequence_tokens": config.max_sequence_tokens,
            "seed": config.seed,
            "adam_beta1": config.beta1,
            "adam_beta2": config.beta2,
            "adam_eps": config.eps,
            "weight_decay": config.weight_decay,
            "grad_clip_norm": config.grad_clip_norm,
            "state_ttl_seconds": config.state_ttl_seconds,
            "sampler_ttl_seconds": None,
            "pipeline_submit_ahead": 1,
        },
    }


def _safe_write_json(path: Path, payload: Mapping[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(f".{path.name}.tmp")
    temporary.write_text(json.dumps(payload, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    os.replace(temporary, path)


def _checkpoint_path(result: Any, label: str) -> str:
    path = getattr(result, "path", None)
    if not isinstance(path, str) or not path.startswith("tinker://"):
        raise TrainingExecutionError(f"{label} did not return a Tinker checkpoint path")
    return path


def _default_nll_calculator() -> Callable[[list[Any], list[Any]], float]:
    try:
        from tinker_cookbook.supervised.common import compute_mean_nll
    except ImportError as exc:
        raise TrainingInputError(
            "native training dependencies are missing; install the native-tinker extra"
        ) from exc
    return compute_mean_nll


def _mean_nll(
    result: Any,
    data: Sequence[Any],
    calculate: Callable[[list[Any], list[Any]], float],
    label: str,
) -> float:
    try:
        outputs = result.loss_fn_outputs
        if len(outputs) != len(data):
            raise ValueError("output count")
        logprobs = [output["logprobs"] for output in outputs]
        weights = [datum.loss_fn_inputs["weights"] for datum in data]
        value = float(calculate(logprobs, weights))
    except (AttributeError, KeyError, TypeError, ValueError) as exc:
        raise TrainingExecutionError(f"{label} returned invalid loss outputs") from exc
    if not math.isfinite(value):
        raise TrainingExecutionError(f"{label} returned a non-finite NLL")
    return value


def _numeric_metrics(result: Any) -> dict[str, float]:
    metrics = getattr(result, "metrics", None)
    if not isinstance(metrics, Mapping):
        return {}
    safe: dict[str, float] = {}
    for key, value in metrics.items():
        if (
            isinstance(key, str)
            and _CHECKPOINT_NAME_RE.fullmatch(key) is not None
            and isinstance(value, (int, float))
            and not isinstance(value, bool)
            and math.isfinite(value)
        ):
            safe[key] = float(value)
    return safe


def execute_training(
    config: TrainConfig,
    prepared: PreparedCorpus,
    plan: Mapping[str, Any],
    *,
    service_factory: Callable[..., Any] | None = None,
    adam_factory: Callable[..., Any] | None = None,
    nll_calculator: Callable[[list[Any], list[Any]], float] | None = None,
    environ: Mapping[str, str] | None = None,
) -> dict[str, Any]:
    """Execute one bounded paid run; validation examples are never passed to training."""

    os.environ[TINKER_TELEMETRY_ENV_VAR] = TINKER_TELEMETRY_VALUE
    if not config.execute:
        raise TrainingInputError("paid training requires --execute")
    if not config.project_id:
        raise TrainingInputError("--project-id is required with --execute")
    if config.run_dir is None:
        raise TrainingInputError("--run-dir is required with --execute")
    batches = _planned_batches(config, prepared.train_examples)
    if not batches:
        raise TrainingInputError("training partition produced no batches")
    validation_data = [item.datum for item in prepared.validation_examples]
    if not validation_data:
        raise TrainingInputError("validation partition is empty")
    source = os.environ if environ is None else environ
    if not source.get(KEY_ENV_VAR, "").strip():
        raise TrainingInputError(f"{KEY_ENV_VAR} is required with --execute")
    if config.run_dir.exists() or config.run_dir.is_symlink():
        raise TrainingInputError("--run-dir must be a new path for each paid run")
    run_dir = config.run_dir.resolve(strict=False)
    try:
        run_dir.mkdir(parents=True, exist_ok=False)
    except FileExistsError as exc:  # pragma: no cover - protects against a creation race
        raise TrainingInputError("--run-dir must be a new path for each paid run") from exc
    metadata_path = run_dir / "run_metadata.json"
    checkpoints_path = run_dir / "checkpoints.json"
    metadata: dict[str, Any] = {
        **dict(plan),
        "status": "initialising",
        "started_at": datetime.now(UTC).isoformat(),
        "completed_steps": 0,
        "training_steps": [],
    }
    _safe_write_json(metadata_path, metadata)

    try:
        if service_factory is None or adam_factory is None:
            import tinker

            service_factory = service_factory or tinker.ServiceClient
            adam_factory = adam_factory or tinker.AdamParams
        calculate_nll = nll_calculator or _default_nll_calculator()
        service = service_factory(
            project_id=config.project_id,
            user_metadata={
                "recipe": CORPUS_BUILDER,
                "corpus_sha256": prepared.corpus.corpus_sha256,
            },
        )
        training_client = service.create_lora_training_client(
            base_model=MODEL_ID,
            rank=config.rank,
            seed=config.seed,
            train_mlp=True,
            train_attn=True,
            train_unembed=True,
            user_metadata={"partition": "development-train"},
        )
        adam = adam_factory(
            learning_rate=config.learning_rate,
            beta1=config.beta1,
            beta2=config.beta2,
            eps=config.eps,
            weight_decay=config.weight_decay,
            grad_clip_norm=config.grad_clip_norm,
        )
        validation_before_result = training_client.forward(validation_data, LOSS_FUNCTION).result()
        metadata["validation_nll_before"] = _mean_nll(
            validation_before_result,
            validation_data,
            calculate_nll,
            "pre-training validation",
        )
        metadata["status"] = "training"
        _safe_write_json(metadata_path, metadata)

        current_batch = batches[0]
        current_data = [item.datum for item in current_batch]
        current_forward = training_client.forward_backward(current_data, LOSS_FUNCTION)
        current_optim = training_client.optim_step(adam)
        for step, next_batch in enumerate(batches[1:], start=1):
            # One-step lookahead: submit the next forward/backward before waiting
            # for the current optimiser, matching Tinker's pipelining guidance.
            next_data = [item.datum for item in next_batch]
            next_forward = training_client.forward_backward(next_data, LOSS_FUNCTION)
            forward_result = current_forward.result()
            optim_result = current_optim.result()
            metadata["training_steps"].append(
                {
                    "assistant_completion_tokens": sum(
                        item.completion_tokens for item in current_batch
                    ),
                    "forward_metrics": _numeric_metrics(forward_result),
                    "num_sequences": len(current_batch),
                    "optim_metrics": _numeric_metrics(optim_result),
                    "step": step,
                    "total_tokens": sum(item.total_tokens for item in current_batch),
                    "train_nll": _mean_nll(
                        forward_result,
                        current_data,
                        calculate_nll,
                        f"training step {step}",
                    ),
                }
            )
            metadata["completed_steps"] = step
            _safe_write_json(metadata_path, metadata)
            current_optim = training_client.optim_step(adam)
            current_forward = next_forward
            current_batch = next_batch
            current_data = next_data
        forward_result = current_forward.result()
        optim_result = current_optim.result()
        final_step = len(batches)
        metadata["training_steps"].append(
            {
                "assistant_completion_tokens": sum(
                    item.completion_tokens for item in current_batch
                ),
                "forward_metrics": _numeric_metrics(forward_result),
                "num_sequences": len(current_batch),
                "optim_metrics": _numeric_metrics(optim_result),
                "step": final_step,
                "total_tokens": sum(item.total_tokens for item in current_batch),
                "train_nll": _mean_nll(
                    forward_result,
                    current_data,
                    calculate_nll,
                    f"training step {final_step}",
                ),
            }
        )
        metadata["completed_steps"] = len(batches)

        validation_after_result = training_client.forward(validation_data, LOSS_FUNCTION).result()
        metadata["validation_nll_after"] = _mean_nll(
            validation_after_result,
            validation_data,
            calculate_nll,
            "post-training validation",
        )
        metadata["validation_nll_change"] = (
            metadata["validation_nll_after"] - metadata["validation_nll_before"]
        )
        _safe_write_json(metadata_path, metadata)

        state_future = training_client.save_state(
            f"{config.checkpoint_prefix}-state",
            ttl_seconds=config.state_ttl_seconds,
            overwrite=False,
            user_metadata={"corpus_sha256": prepared.corpus.corpus_sha256},
        )
        sampler_future = training_client.save_weights_for_sampler(
            f"{config.checkpoint_prefix}-sampler",
            ttl_seconds=None,
            user_metadata={"corpus_sha256": prepared.corpus.corpus_sha256},
        )
        state_path = _checkpoint_path(state_future.result(), "training state")
        sampler_path = _checkpoint_path(sampler_future.result(), "sampler weights")
        checkpoints = {
            "schema_version": 1,
            "base_model": MODEL_ID,
            "corpus_sha256": prepared.corpus.corpus_sha256,
            "state": {"path": state_path, "ttl_seconds": config.state_ttl_seconds},
            "sampler": {"path": sampler_path, "ttl_seconds": None},
        }
        _safe_write_json(checkpoints_path, checkpoints)
        metadata.update(
            status="complete",
            completed_at=datetime.now(UTC).isoformat(),
            state_checkpoint=state_path,
            sampler_checkpoint=sampler_path,
        )
        _safe_write_json(metadata_path, metadata)
        return checkpoints
    except TrainingInputError:
        raise
    except Exception as exc:
        metadata.update(
            status="failed",
            completed_at=datetime.now(UTC).isoformat(),
            error_type=type(exc).__name__,
        )
        _safe_write_json(metadata_path, metadata)
        raise TrainingExecutionError("Tinker training failed; inspect the named session") from None


def _bounded_float(text: str, *, minimum: float, maximum: float, label: str) -> float:
    try:
        value = float(text)
    except ValueError as exc:
        raise argparse.ArgumentTypeError(f"{label} must be numeric") from exc
    if not math.isfinite(value) or not minimum <= value <= maximum:
        raise argparse.ArgumentTypeError(f"{label} must be between {minimum} and {maximum}")
    return value


def _bounded_int(text: str, *, minimum: int, maximum: int, label: str) -> int:
    try:
        value = int(text)
    except ValueError as exc:
        raise argparse.ArgumentTypeError(f"{label} must be an integer") from exc
    if not minimum <= value <= maximum:
        raise argparse.ArgumentTypeError(f"{label} must be between {minimum} and {maximum}")
    return value


def parse_args(argv: Sequence[str] | None = None) -> TrainConfig:
    parser = argparse.ArgumentParser(
        description="Validate and optionally execute bounded Qwen3.8-27B LoRA SFT."
    )
    parser.add_argument("--manifest", type=Path, default=DEFAULT_CORPUS_MANIFEST)
    parser.add_argument("--split-manifest", type=Path, default=DEFAULT_SPLIT_MANIFEST)
    parser.add_argument("--execute", action="store_true", help="authorise paid Tinker training")
    parser.add_argument("--project-id", help="explicit writable Tinker project ID")
    parser.add_argument(
        "--run-dir", type=Path, help="local metadata directory (required to execute)"
    )
    parser.add_argument("--rank", type=int, choices=(8, 16, 32, 64, 128), default=32)
    parser.add_argument(
        "--learning-rate",
        type=lambda value: _bounded_float(value, minimum=1e-7, maximum=1e-2, label="learning rate"),
        default=1e-4,
    )
    parser.add_argument(
        "--batch-size",
        type=lambda value: _bounded_int(
            value, minimum=1, maximum=MAX_BATCH_SIZE, label="batch size"
        ),
        default=4,
    )
    parser.add_argument(
        "--epochs",
        type=lambda value: _bounded_int(value, minimum=1, maximum=MAX_EPOCHS, label="epochs"),
        default=1,
    )
    parser.add_argument(
        "--max-steps",
        type=lambda value: _bounded_int(value, minimum=1, maximum=MAX_STEPS, label="max steps"),
        default=20,
    )
    parser.add_argument(
        "--max-sequence-tokens",
        type=lambda value: _bounded_int(
            value, minimum=1024, maximum=65536, label="max sequence tokens"
        ),
        default=65536,
    )
    parser.add_argument("--seed", type=int, default=2026)
    parser.add_argument(
        "--beta1",
        type=lambda value: _bounded_float(value, minimum=0.0, maximum=0.9999, label="beta1"),
        default=0.9,
    )
    parser.add_argument(
        "--beta2",
        type=lambda value: _bounded_float(value, minimum=0.0, maximum=0.99999, label="beta2"),
        default=0.95,
    )
    parser.add_argument(
        "--eps",
        type=lambda value: _bounded_float(value, minimum=1e-12, maximum=1e-3, label="eps"),
        default=1e-8,
    )
    parser.add_argument(
        "--weight-decay",
        type=lambda value: _bounded_float(value, minimum=0.0, maximum=1.0, label="weight decay"),
        default=0.0,
    )
    parser.add_argument(
        "--grad-clip-norm",
        type=lambda value: _bounded_float(
            value, minimum=0.0, maximum=100.0, label="gradient clip norm"
        ),
        default=0.0,
    )
    parser.add_argument(
        "--state-ttl-seconds",
        type=lambda value: _bounded_int(
            value, minimum=3600, maximum=2_592_000, label="state TTL seconds"
        ),
        default=604800,
    )
    parser.add_argument("--checkpoint-prefix", default="formulabench-sft")
    parser.add_argument(
        "--allow-tokenizer-download",
        action="store_true",
        help="allow a Hugging Face tokenizer download during local preflight",
    )
    args = parser.parse_args(argv)
    if _CHECKPOINT_NAME_RE.fullmatch(args.checkpoint_prefix) is None:
        parser.error(
            "--checkpoint-prefix must contain only letters, numbers, dot, underscore or dash"
        )
    if args.execute and not args.project_id:
        parser.error("--project-id is required with --execute")
    if args.execute and args.run_dir is None:
        parser.error("--run-dir is required with --execute")
    return TrainConfig(
        manifest_path=args.manifest,
        split_manifest_path=args.split_manifest,
        execute=args.execute,
        project_id=args.project_id,
        run_dir=args.run_dir,
        rank=args.rank,
        learning_rate=args.learning_rate,
        batch_size=args.batch_size,
        epochs=args.epochs,
        max_steps=args.max_steps,
        max_sequence_tokens=args.max_sequence_tokens,
        seed=args.seed,
        beta1=args.beta1,
        beta2=args.beta2,
        eps=args.eps,
        weight_decay=args.weight_decay,
        grad_clip_norm=args.grad_clip_norm,
        state_ttl_seconds=args.state_ttl_seconds,
        checkpoint_prefix=args.checkpoint_prefix,
        allow_tokenizer_download=args.allow_tokenizer_download,
    )


def _check_runtime_versions() -> None:
    expected = {"tinker": "0.27.1", "tinker-cookbook": "0.5.7"}
    for package, expected_version in expected.items():
        try:
            installed = version(package)
        except PackageNotFoundError as exc:
            raise TrainingInputError(
                "native training dependencies are missing; install the native-tinker extra"
            ) from exc
        if installed != expected_version:
            raise TrainingInputError(
                f"{package} must be pinned to {expected_version}; found {installed}"
            )


def main(argv: Sequence[str] | None = None) -> int:
    # The runner never reads or prints credential values. The SDK reads its key
    # internally only after --execute passes every local preflight.
    os.environ[TINKER_TELEMETRY_ENV_VAR] = TINKER_TELEMETRY_VALUE
    try:
        config = parse_args(argv)
        _check_runtime_versions()
        corpus = load_validated_corpus(config.manifest_path, config.split_manifest_path)
        prepared = prepare_corpus(
            corpus,
            max_sequence_tokens=config.max_sequence_tokens,
            allow_tokenizer_download=config.allow_tokenizer_download,
        )
        plan = build_plan(config, prepared)
        print(json.dumps(plan, indent=2, sort_keys=True), flush=True)
        if not config.execute:
            print("Dry run complete: zero Tinker provider calls were made.", flush=True)
            return 0
        checkpoints = execute_training(config, prepared, plan)
        sampler_path = checkpoints["sampler"]["path"]
        print(f"SAMPLER_CHECKPOINT={sampler_path}", flush=True)
        print(
            "Use this value with ./scripts/evaluate_checkpoint.sh --execute.",
            flush=True,
        )
        return 0
    except TrainingInputError as exc:
        print(f"preflight failed: {exc}", file=sys.stderr, flush=True)
        return 2
    except TrainingExecutionError as exc:
        print(str(exc), file=sys.stderr, flush=True)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
