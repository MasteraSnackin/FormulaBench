"""Create a hash-bound base-versus-checkpoint evaluator comparison."""

from __future__ import annotations

import argparse
import hashlib
import json
import math
import os
import re
import tempfile
from collections.abc import Mapping, Sequence
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from formulabench.artifacts import ArtifactError, normalise_task_id
from formulabench.constants import MODEL_ID, MODEL_PROVENANCE
from formulabench.contract import FailureCode
from formulabench.provider import (
    ProviderConfigurationError,
    sampler_checkpoint_provenance,
    validate_sampler_checkpoint,
)
from formulabench.runner import _request_trace_metadata
from training.train_lora import (
    DEFAULT_CORPUS_MANIFEST,
    DEFAULT_SPLIT_MANIFEST,
    TrainingInputError,
    load_validated_corpus,
)

MAX_RESULTS_BYTES = 32 * 1024 * 1024
MAX_PREDICTIONS_BYTES = 32 * 1024 * 1024
MAX_TRACE_BYTES = 32 * 1024 * 1024
MAX_OUTPUT_BYTES = 1024 * 1024 * 1024
METRICS = ("pass_rate", "cell_accuracy", "pass_rate_cell_level", "pass_rate_sheet_level")
_FAILURE_STATUS = re.compile(r"error:[a-z0-9_]+(?:,[a-z0-9_]+)*\Z")
_CHECKPOINT_PROVENANCE = re.compile(r"sampler_checkpoint_sha256:[0-9a-f]{64}\Z")
_SHA256 = re.compile(r"[0-9a-f]{64}\Z")
_FAILURE_CODES = frozenset(code.value for code in FailureCode)
_TRACE_REQUIRED_FIELDS = frozenset(
    {
        "error",
        "failure_codes",
        "input_tokens",
        "latency_ms",
        "model",
        "model_provenance",
        "output_tokens",
        "parse_termination",
        "prompt",
        "renderer",
        "request",
        "response",
        "response_format",
        "response_rejection",
        "step",
        "stop_reason",
    }
)


class ComparisonError(ValueError):
    """An evaluator result cannot support a controlled comparison."""


@dataclass(frozen=True, slots=True)
class TraceEvidence:
    file_hashes: dict[str, str]
    manifest_sha256: str
    prompt_sha256: dict[str, str | None]
    context_input_sha256: dict[str, str | None]
    empty_pre_call_task_ids: tuple[str, ...]


def _sha256(raw: bytes) -> str:
    return hashlib.sha256(raw).hexdigest()


def _canonical_bytes(value: object) -> bytes:
    return json.dumps(
        value,
        ensure_ascii=False,
        separators=(",", ":"),
        sort_keys=True,
    ).encode("utf-8")


def _read_regular(
    path: Path,
    label: str,
    maximum: int,
    *,
    allow_empty: bool = False,
) -> bytes:
    if path.is_symlink() or not path.is_file():
        raise ComparisonError(f"{label} must be a regular non-symlink file")
    try:
        size = path.stat().st_size
        if size < 0 or (size == 0 and not allow_empty) or size > maximum:
            raise ComparisonError(f"{label} has an invalid size")
        raw = path.read_bytes()
        if len(raw) != size:
            raise ComparisonError(f"{label} changed while it was being read")
        return raw
    except OSError as exc:
        raise ComparisonError(f"{label} is not readable") from exc


def _hash_regular(path: Path, label: str, maximum: int) -> str:
    if path.is_symlink() or not path.is_file():
        raise ComparisonError(f"{label} must be a regular non-symlink file")
    try:
        size = path.stat().st_size
        if size <= 0 or size > maximum:
            raise ComparisonError(f"{label} has an invalid size")
        digest = hashlib.sha256()
        with path.open("rb") as handle:
            for chunk in iter(lambda: handle.read(1024 * 1024), b""):
                digest.update(chunk)
        return digest.hexdigest()
    except OSError as exc:
        raise ComparisonError(f"{label} is not readable") from exc


def _load_result(path: Path, label: str) -> tuple[Mapping[str, Any], bytes]:
    raw = _read_regular(path, f"{label} result", MAX_RESULTS_BYTES)
    try:
        payload = json.loads(raw.decode("utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise ComparisonError(f"{label} result is not valid UTF-8 JSON") from exc
    if not isinstance(payload, Mapping):
        raise ComparisonError(f"{label} result root must be an object")
    summary = payload.get("summary")
    items = payload.get("items")
    if not isinstance(summary, Mapping) or not isinstance(items, list) or not items:
        raise ComparisonError(f"{label} result lacks a summary or item records")
    for metric in METRICS:
        value = summary.get(metric)
        if value is not None and (
            isinstance(value, bool)
            or not isinstance(value, (int, float))
            or not math.isfinite(value)
        ):
            raise ComparisonError(f"{label} result has an invalid {metric}")
    return payload, raw


def _safe_task_id(value: object, label: str) -> str:
    if not isinstance(value, str):
        raise ComparisonError(f"{label} has an invalid task ID")
    try:
        return normalise_task_id(value)
    except ArtifactError as exc:
        raise ComparisonError(f"{label} has an invalid task ID") from exc


def _rate(rows: Sequence[Mapping[str, Any]]) -> float | None:
    if not rows:
        return None
    return round(sum(item.get("pass", False) is True for item in rows) / len(rows), 4)


def _strict_result_summary(
    payload: Mapping[str, Any],
    label: str,
) -> tuple[tuple[str, ...], int]:
    ids: list[str] = []
    graded: list[Mapping[str, Any]] = []
    cell_items: list[Mapping[str, Any]] = []
    sheet_items: list[Mapping[str, Any]] = []
    error_count = 0
    for index, raw_item in enumerate(payload["items"]):
        if not isinstance(raw_item, Mapping):
            raise ComparisonError(f"{label} item {index} is not an object")
        task_id = _safe_task_id(raw_item.get("id"), f"{label} item {index}")
        ids.append(task_id)
        instruction_type = raw_item.get("type")
        if not isinstance(instruction_type, str) or not instruction_type:
            raise ComparisonError(f"{label} item {task_id} has an invalid instruction type")
        if instruction_type.startswith("Cell"):
            cell_items.append(raw_item)
        elif instruction_type.startswith("Sheet"):
            sheet_items.append(raw_item)
        else:
            raise ComparisonError(f"{label} item {task_id} has an unknown instruction type")
        status = raw_item.get("status")
        if status == "graded":
            cells = raw_item.get("cells")
            correct = raw_item.get("correct")
            passed = raw_item.get("pass")
            mismatches = raw_item.get("mismatches")
            if (
                isinstance(cells, bool)
                or not isinstance(cells, int)
                or cells <= 0
                or isinstance(correct, bool)
                or not isinstance(correct, int)
                or not 0 <= correct <= cells
                or not isinstance(passed, bool)
                or passed != (correct == cells)
                or not isinstance(mismatches, list)
                or (passed and bool(mismatches))
                or (not passed and not mismatches)
            ):
                raise ComparisonError(f"{label} item {task_id} has invalid graded fields")
            graded.append(raw_item)
        elif status == "error":
            error = raw_item.get("error")
            if not isinstance(error, str) or any(
                field in raw_item for field in ("cells", "correct", "pass")
            ):
                raise ComparisonError(f"{label} item {task_id} has invalid error fields")
            error_count += 1
        else:
            raise ComparisonError(f"{label} item {task_id} has an unsupported evaluator status")
    if len(ids) != len(set(ids)):
        raise ComparisonError(f"{label} result contains duplicate task IDs")

    summary = payload["summary"]
    cell_denominator = sum(int(item["cells"]) for item in graded)
    expected_summary: dict[str, int | float | None] = {
        "items": len(ids),
        "graded": len(graded),
        "missing": 0,
        "errors": error_count,
        "pass_rate": round(
            sum(item.get("pass", False) is True for item in payload["items"]) / len(ids),
            4,
        ),
        "cell_accuracy": (
            round(sum(int(item["correct"]) for item in graded) / cell_denominator, 4)
            if cell_denominator
            else None
        ),
        "pass_rate_cell_level": _rate(cell_items),
        "pass_rate_sheet_level": _rate(sheet_items),
    }
    for field, expected in expected_summary.items():
        reported = summary.get(field)
        if field in {"items", "graded", "missing", "errors"} and (
            isinstance(reported, bool) or not isinstance(reported, int)
        ):
            raise ComparisonError(f"{label} result summary has an invalid {field}")
        if reported != expected:
            raise ComparisonError(f"{label} result summary has an invalid {field}")
    return tuple(ids), error_count


def _basic_task_ids(payload: Mapping[str, Any], label: str) -> tuple[str, ...]:
    ids: list[str] = []
    for index, raw_item in enumerate(payload["items"]):
        if not isinstance(raw_item, Mapping):
            raise ComparisonError(f"{label} item {index} is not an object")
        ids.append(_safe_task_id(raw_item.get("id"), f"{label} item {index}"))
    if len(ids) != len(set(ids)):
        raise ComparisonError(f"{label} result contains duplicate task IDs")
    return tuple(ids)


def _validate_cross_arm_result_invariants(
    base: Mapping[str, Any],
    checkpoint: Mapping[str, Any],
) -> None:
    for base_item, checkpoint_item in zip(base["items"], checkpoint["items"], strict=True):
        task_id = str(base_item["id"])
        if base_item["type"] != checkpoint_item["type"]:
            raise ComparisonError(
                f"task {task_id} has different evaluator types across comparison arms"
            )
        if (
            base_item["status"] == "graded"
            and checkpoint_item["status"] == "graded"
            and base_item["cells"] != checkpoint_item["cells"]
        ):
            raise ComparisonError(
                f"task {task_id} has different graded-cell counts across comparison arms"
            )


def _basic_error_count(summary: Mapping[str, Any], label: str) -> int:
    value = summary.get("errors", 0)
    if isinstance(value, bool) or not isinstance(value, int) or value < 0:
        raise ComparisonError(f"{label} result has an invalid evaluator error count")
    return value


def _load_jsonl(
    path: Path,
    label: str,
    maximum: int,
    *,
    allow_empty: bool = False,
) -> tuple[list[Mapping[str, Any]], bytes]:
    raw = _read_regular(path, label, maximum, allow_empty=allow_empty)
    try:
        lines = raw.decode("utf-8").splitlines()
    except UnicodeDecodeError as exc:
        raise ComparisonError(f"{label} is not valid UTF-8 JSONL") from exc
    if not lines and allow_empty:
        return [], raw
    if not lines or any(not line.strip() for line in lines):
        raise ComparisonError(f"{label} contains an empty JSONL record")

    records: list[Mapping[str, Any]] = []
    for index, line in enumerate(lines, start=1):
        try:
            record = json.loads(line)
        except json.JSONDecodeError as exc:
            raise ComparisonError(f"{label} line {index} is not valid JSON") from exc
        if not isinstance(record, Mapping):
            raise ComparisonError(f"{label} line {index} is not an object")
        records.append(record)
    return records, raw


def _load_predictions(path: Path, label: str) -> tuple[list[Mapping[str, Any]], bytes]:
    records, raw = _load_jsonl(path, f"{label} predictions", MAX_PREDICTIONS_BYTES)
    ids: list[str] = []
    for index, record in enumerate(records, start=1):
        task_id = _safe_task_id(record.get("id"), f"{label} prediction line {index}")
        ids.append(task_id)
        if set(record) != {"id", "output", "status", "failure_codes"}:
            raise ComparisonError(f"{label} prediction {task_id} has an invalid schema")
        output = record.get("output")
        if output != f"outputs/{task_id}.xlsx":
            raise ComparisonError(f"{label} prediction {task_id} has an invalid output path")
        status = record.get("status")
        if status != "ok" and (
            not isinstance(status, str) or _FAILURE_STATUS.fullmatch(status) is None
        ):
            raise ComparisonError(f"{label} prediction {task_id} has an invalid status")
        failure_codes = record.get("failure_codes")
        if not isinstance(failure_codes, list) or any(
            not isinstance(code, str) or not code for code in failure_codes
        ):
            raise ComparisonError(f"{label} prediction {task_id} has invalid failure codes")
        if len(failure_codes) != len(set(failure_codes)) or any(
            code not in _FAILURE_CODES for code in failure_codes
        ):
            raise ComparisonError(f"{label} prediction {task_id} has invalid failure codes")
        expected_codes = [] if status == "ok" else status.removeprefix("error:").split(",")
        if failure_codes != expected_codes:
            raise ComparisonError(f"{label} prediction {task_id} status and failure codes disagree")
    if len(ids) != len(set(ids)):
        raise ComparisonError(f"{label} predictions contain duplicate task IDs")
    return records, raw


def _load_traces(
    path: Path,
    label: str,
    task_ids: Sequence[str],
    prediction_statuses: Mapping[str, str],
    expected_model_provenance: str,
) -> TraceEvidence:
    if path.is_symlink() or not path.is_dir():
        raise ComparisonError(f"{label} traces must be a regular non-symlink directory")
    try:
        entries = list(path.iterdir())
    except OSError as exc:
        raise ComparisonError(f"{label} traces directory is not readable") from exc

    expected_names = {f"{task_id}.jsonl" for task_id in task_ids}
    actual_names = {entry.name for entry in entries}
    if actual_names != expected_names:
        raise ComparisonError(f"{label} traces do not exactly match the comparison tasks")

    trace_hashes: dict[str, str] = {}
    prompt_hashes: dict[str, str | None] = {}
    context_input_hashes: dict[str, str | None] = {}
    empty_pre_call_task_ids: list[str] = []
    expected_request = _request_trace_metadata(expected_model_provenance)
    for task_id in task_ids:
        trace_path = path / f"{task_id}.jsonl"
        records, raw = _load_jsonl(
            trace_path,
            f"{label} trace {task_id}",
            MAX_TRACE_BYTES,
            allow_empty=True,
        )
        status = prediction_statuses[task_id]
        expected_failure_codes = [] if status == "ok" else status.removeprefix("error:").split(",")
        if not records:
            if status != "error:context_build_failed":
                raise ComparisonError(
                    f"{label} trace {task_id} is empty without a pre-call context failure"
                )
            empty_pre_call_task_ids.append(task_id)
            trace_hashes[task_id] = _sha256(raw)
            prompt_hashes[task_id] = None
            context_input_hashes[task_id] = None
            continue
        if status == "error:context_build_failed":
            raise ComparisonError(
                f"{label} trace {task_id} contains a model call after a context failure"
            )
        if len(records) != 1:
            raise ComparisonError(f"{label} trace {task_id} must contain exactly one model call")

        for index, record in enumerate(records, start=1):
            if _TRACE_REQUIRED_FIELDS.difference(record):
                raise ComparisonError(
                    f"{label} trace {task_id} line {index} is missing required fields"
                )
            step = record.get("step")
            if isinstance(step, bool) or not isinstance(step, int) or step != index:
                raise ComparisonError(f"{label} trace {task_id} has invalid step order")
            if record.get("model") != MODEL_ID:
                raise ComparisonError(f"{label} trace {task_id} line {index} has wrong model")
            if record.get("model_provenance") != expected_model_provenance:
                raise ComparisonError(
                    f"{label} trace {task_id} line {index} has wrong model provenance"
                )
            if record.get("request") != expected_request:
                raise ComparisonError(
                    f"{label} trace {task_id} line {index} has wrong fixed request settings"
                )
            for field in (
                "prompt",
                "response",
                "error",
                "stop_reason",
                "parse_termination",
                "response_format",
                "response_rejection",
            ):
                if record.get(field) is not None and not isinstance(record.get(field), str):
                    raise ComparisonError(
                        f"{label} trace {task_id} line {index} has an invalid {field}"
                    )
            if not isinstance(record.get("prompt"), str):
                raise ComparisonError(f"{label} trace {task_id} line {index} has an invalid prompt")
            prompt_hashes[task_id] = _sha256(record["prompt"].encode("utf-8"))
            for field in ("input_tokens", "output_tokens", "latency_ms"):
                value = record.get(field)
                if value is not None and (
                    isinstance(value, bool) or not isinstance(value, int) or value < 0
                ):
                    raise ComparisonError(
                        f"{label} trace {task_id} line {index} has an invalid {field}"
                    )
            failure_codes = record.get("failure_codes")
            if not isinstance(failure_codes, list) or any(
                not isinstance(code, str) or not code for code in failure_codes
            ):
                raise ComparisonError(
                    f"{label} trace {task_id} line {index} has invalid failure codes"
                )
            if len(failure_codes) != len(set(failure_codes)) or any(
                code not in _FAILURE_CODES for code in failure_codes
            ):
                raise ComparisonError(
                    f"{label} trace {task_id} line {index} has invalid failure codes"
                )
            if record.get("renderer") != expected_request["renderer"]:
                raise ComparisonError(f"{label} trace {task_id} line {index} has wrong renderer")
            context = record.get("context")
            if status == "ok" and not isinstance(context, Mapping):
                raise ComparisonError(
                    f"{label} trace {task_id} line {index} lacks workbook context"
                )
            if isinstance(context, Mapping):
                if set(context) != {
                    "characters",
                    "included_cells",
                    "input_sha256",
                    "omitted_cells",
                    "truncated",
                }:
                    raise ComparisonError(
                        f"{label} trace {task_id} line {index} has invalid context fields"
                    )
                for field in ("characters", "included_cells", "omitted_cells"):
                    value = context.get(field)
                    if isinstance(value, bool) or not isinstance(value, int) or value < 0:
                        raise ComparisonError(
                            f"{label} trace {task_id} line {index} has invalid context fields"
                        )
                if not isinstance(context.get("truncated"), bool):
                    raise ComparisonError(
                        f"{label} trace {task_id} line {index} has invalid context fields"
                    )
            context_input_sha256 = (
                context.get("input_sha256") if isinstance(context, Mapping) else None
            )
            if isinstance(context, Mapping) and (
                not isinstance(context_input_sha256, str)
                or _SHA256.fullmatch(context_input_sha256) is None
            ):
                raise ComparisonError(
                    f"{label} trace {task_id} line {index} has invalid context provenance"
                )
            context_input_hashes[task_id] = context_input_sha256

        final_record = records[-1]
        if final_record["failure_codes"] != expected_failure_codes:
            raise ComparisonError(f"{label} prediction {task_id} and trace failure codes disagree")
        if status == "ok" and not (
            final_record.get("error") is None and isinstance(final_record.get("response"), str)
        ):
            raise ComparisonError(
                f"{label} prediction {task_id} is accepted without a successful trace"
            )
        if status != "ok" and not isinstance(final_record.get("error"), str):
            raise ComparisonError(f"{label} failed prediction {task_id} lacks a failed trace")
        trace_hashes[task_id] = _sha256(raw)

    trace_manifest = [{"task_id": task_id, "sha256": trace_hashes[task_id]} for task_id in task_ids]
    return TraceEvidence(
        file_hashes=trace_hashes,
        manifest_sha256=_sha256(_canonical_bytes(trace_manifest)),
        prompt_sha256=prompt_hashes,
        context_input_sha256=context_input_hashes,
        empty_pre_call_task_ids=tuple(empty_pre_call_task_ids),
    )


def _load_outputs(
    predictions_path: Path,
    label: str,
    task_ids: Sequence[str],
) -> tuple[dict[str, str], str]:
    outputs_path = predictions_path.parent / "outputs"
    if outputs_path.is_symlink() or not outputs_path.is_dir():
        raise ComparisonError(f"{label} outputs must be a regular non-symlink directory")
    try:
        entries = list(outputs_path.iterdir())
    except OSError as exc:
        raise ComparisonError(f"{label} outputs directory is not readable") from exc
    expected_names = {f"{task_id}.xlsx" for task_id in task_ids}
    actual_names = {entry.name for entry in entries}
    if actual_names != expected_names:
        raise ComparisonError(f"{label} outputs do not exactly match the comparison tasks")

    output_hashes = {
        task_id: _hash_regular(
            outputs_path / f"{task_id}.xlsx",
            f"{label} output {task_id}",
            MAX_OUTPUT_BYTES,
        )
        for task_id in task_ids
    }
    output_manifest = [
        {"task_id": task_id, "sha256": output_hashes[task_id]} for task_id in task_ids
    ]
    return output_hashes, _sha256(_canonical_bytes(output_manifest))


def _arm_evidence(
    *,
    label: str,
    predictions_path: Path,
    traces_path: Path,
    expected_task_ids: Sequence[str],
    expected_model_provenance: str,
    expected_prompt_sha256: Mapping[str, str] | None = None,
) -> dict[str, Any]:
    predictions, predictions_raw = _load_predictions(predictions_path, label)
    try:
        canonical_traces = (predictions_path.parent / "traces").resolve(strict=True)
        supplied_traces = traces_path.resolve(strict=True)
    except (OSError, RuntimeError) as exc:
        raise ComparisonError(f"{label} trace path cannot be resolved") from exc
    if supplied_traces != canonical_traces:
        raise ComparisonError(f"{label} traces must be beside its predictions file")
    prediction_ids = tuple(
        _safe_task_id(record.get("id"), f"{label} prediction") for record in predictions
    )
    if prediction_ids != tuple(expected_task_ids):
        raise ComparisonError(
            f"{label} predictions do not match the exact ordered comparison tasks"
        )

    statuses = {str(record["id"]): str(record["status"]) for record in predictions}
    accepted = sum(status == "ok" for status in statuses.values())
    failed = len(statuses) - accepted
    if accepted == 0:
        raise ComparisonError(f"{label} arm has zero accepted predictions")

    trace_evidence = _load_traces(
        traces_path,
        label,
        prediction_ids,
        statuses,
        expected_model_provenance,
    )
    if expected_prompt_sha256 is not None:
        if set(expected_prompt_sha256) != set(prediction_ids):
            raise ComparisonError(f"{label} expected prompt bindings do not match its tasks")
        for task_id in prediction_ids:
            expected_prompt = expected_prompt_sha256[task_id]
            if not isinstance(expected_prompt, str) or _SHA256.fullmatch(expected_prompt) is None:
                raise ComparisonError(f"{label} expected prompt binding is invalid")
            observed = trace_evidence.prompt_sha256[task_id]
            if observed is not None and observed != expected_prompt:
                raise ComparisonError(
                    f"{label} trace {task_id} does not match the validated corpus prompt"
                )
    output_hashes, outputs_manifest_hash = _load_outputs(
        predictions_path,
        label,
        prediction_ids,
    )
    bindings = [
        {
            "model_provenance": expected_model_provenance,
            "output_sha256": output_hashes[task_id],
            "prediction_sha256": _sha256(_canonical_bytes(prediction)),
            "prompt_sha256": trace_evidence.prompt_sha256[task_id],
            "task_id": task_id,
            "trace_sha256": trace_evidence.file_hashes[task_id],
        }
        for task_id, prediction in zip(prediction_ids, predictions, strict=True)
    ]
    return {
        "model_provenance": expected_model_provenance,
        "prediction_counts": {
            "accepted": accepted,
            "empty_pre_call_failures": len(trace_evidence.empty_pre_call_task_ids),
            "failed": failed,
            "total": len(predictions),
        },
        "output_files_sha256": output_hashes,
        "outputs_manifest_sha256": outputs_manifest_hash,
        "predictions_sha256": _sha256(predictions_raw),
        "trace_files_sha256": trace_evidence.file_hashes,
        "traces_manifest_sha256": trace_evidence.manifest_sha256,
        "trace_prompt_sha256": trace_evidence.prompt_sha256,
        "trace_context_input_sha256": trace_evidence.context_input_sha256,
        "empty_pre_call_task_ids": list(trace_evidence.empty_pre_call_task_ids),
        "prediction_trace_bindings_sha256": _sha256(_canonical_bytes(bindings)),
    }


def compare_results(
    base_path: Path,
    checkpoint_path: Path,
    *,
    base_predictions_path: Path | None = None,
    base_traces_path: Path | None = None,
    checkpoint_predictions_path: Path | None = None,
    checkpoint_traces_path: Path | None = None,
    checkpoint_model_provenance: str | None = None,
    expected_task_ids: Sequence[str] | None = None,
    expected_prompt_sha256: Mapping[str, str] | None = None,
) -> dict[str, Any]:
    evidence_inputs = (
        base_predictions_path,
        base_traces_path,
        checkpoint_predictions_path,
        checkpoint_traces_path,
        checkpoint_model_provenance,
    )
    evidence_supplied = all(value is not None for value in evidence_inputs)
    if any(value is not None for value in evidence_inputs) and not evidence_supplied:
        raise ComparisonError("comparison evidence arguments must be supplied together")
    if expected_prompt_sha256 is not None and not evidence_supplied:
        raise ComparisonError("prompt bindings require complete comparison evidence")

    base, base_raw = _load_result(base_path, "base")
    checkpoint, checkpoint_raw = _load_result(checkpoint_path, "checkpoint")
    if evidence_supplied:
        base_ids, base_errors = _strict_result_summary(base, "base")
        checkpoint_ids, checkpoint_errors = _strict_result_summary(checkpoint, "checkpoint")
    else:
        base_ids = _basic_task_ids(base, "base")
        checkpoint_ids = _basic_task_ids(checkpoint, "checkpoint")
        base_errors = _basic_error_count(base["summary"], "base")
        checkpoint_errors = _basic_error_count(checkpoint["summary"], "checkpoint")
    if base_ids != checkpoint_ids:
        raise ComparisonError("base and checkpoint results do not contain the same ordered tasks")
    if expected_task_ids is not None and base_ids != tuple(expected_task_ids):
        raise ComparisonError("comparison tasks do not match the validated corpus partition")
    if evidence_supplied:
        _validate_cross_arm_result_invariants(base, checkpoint)

    evidence: dict[str, Any] | None = None
    if evidence_supplied:
        assert base_predictions_path is not None
        assert base_traces_path is not None
        assert checkpoint_predictions_path is not None
        assert checkpoint_traces_path is not None
        assert checkpoint_model_provenance is not None
        if _CHECKPOINT_PROVENANCE.fullmatch(checkpoint_model_provenance) is None:
            raise ComparisonError("checkpoint model provenance is not an exact SHA-256 binding")
        evidence = {
            "base": _arm_evidence(
                label="base",
                predictions_path=base_predictions_path,
                traces_path=base_traces_path,
                expected_task_ids=base_ids,
                expected_model_provenance=MODEL_PROVENANCE,
                expected_prompt_sha256=expected_prompt_sha256,
            ),
            "checkpoint": _arm_evidence(
                label="checkpoint",
                predictions_path=checkpoint_predictions_path,
                traces_path=checkpoint_traces_path,
                expected_task_ids=checkpoint_ids,
                expected_model_provenance=checkpoint_model_provenance,
                expected_prompt_sha256=expected_prompt_sha256,
            ),
        }
        if evidence["base"]["trace_prompt_sha256"] != evidence["checkpoint"]["trace_prompt_sha256"]:
            raise ComparisonError("base and checkpoint traces used different prompts")
        if (
            evidence["base"]["empty_pre_call_task_ids"]
            != evidence["checkpoint"]["empty_pre_call_task_ids"]
        ):
            raise ComparisonError(
                "base and checkpoint traces have different pre-call failure tasks"
            )
        for task_id in base_ids:
            base_context = evidence["base"]["trace_context_input_sha256"][task_id]
            checkpoint_context = evidence["checkpoint"]["trace_context_input_sha256"][task_id]
            if (
                base_context is not None
                and checkpoint_context is not None
                and base_context != checkpoint_context
            ):
                raise ComparisonError(
                    f"task {task_id} has different workbook context across comparison arms"
                )

    base_summary = dict(base["summary"])
    checkpoint_summary = dict(checkpoint["summary"])
    cell_accuracy_comparable = base_errors == 0 and checkpoint_errors == 0
    comparability_reason = None
    if not cell_accuracy_comparable:
        comparability_reason = (
            "cell_accuracy excludes evaluator-error tasks from its cell denominator, so "
            "the checkpoint-minus-base delta is not comparable when either arm has errors"
        )

    deltas: dict[str, float | None] = {}
    for metric in METRICS:
        base_value = base_summary.get(metric)
        checkpoint_value = checkpoint_summary.get(metric)
        deltas[metric] = (
            round(float(checkpoint_value) - float(base_value), 4)
            if base_value is not None and checkpoint_value is not None
            else None
        )
    if not cell_accuracy_comparable:
        deltas["cell_accuracy"] = None

    result = {
        "schema_version": 2,
        "task_count": len(base_ids),
        "ordered_task_ids": list(base_ids),
        "base_results_sha256": _sha256(base_raw),
        "checkpoint_results_sha256": _sha256(checkpoint_raw),
        "evidence_validation": "provenance-and-hash-bound" if evidence else "not-supplied",
        "evaluator_errors": {
            "base": base_errors,
            "checkpoint": checkpoint_errors,
            "checkpoint_minus_base": checkpoint_errors - base_errors,
            "counts_equal": base_errors == checkpoint_errors,
        },
        "metric_comparability": {
            "cell_accuracy": {
                "comparable": cell_accuracy_comparable,
                "reason": comparability_reason,
            }
        },
        "base": base_summary,
        "checkpoint": checkpoint_summary,
        "checkpoint_minus_base": deltas,
    }
    if evidence is not None:
        result["evidence"] = evidence
    return result


def _write_new_json(path: Path, payload: Mapping[str, Any]) -> None:
    if path.exists() or path.is_symlink():
        raise ComparisonError("comparison output must be a new path")
    path.parent.mkdir(parents=True, exist_ok=True)
    descriptor, temporary_name = tempfile.mkstemp(
        dir=path.parent,
        prefix=f".{path.name}.",
        suffix=".tmp",
    )
    temporary = Path(temporary_name)
    try:
        with os.fdopen(descriptor, "w", encoding="utf-8") as handle:
            json.dump(payload, handle, ensure_ascii=False, indent=2, sort_keys=True)
            handle.write("\n")
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(temporary, path)
    finally:
        temporary.unlink(missing_ok=True)


def parse_args(argv: Sequence[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--base", required=True, type=Path)
    parser.add_argument("--base-predictions", required=True, type=Path)
    parser.add_argument("--base-traces", required=True, type=Path)
    parser.add_argument("--checkpoint", required=True, type=Path)
    parser.add_argument("--checkpoint-predictions", required=True, type=Path)
    parser.add_argument("--checkpoint-traces", required=True, type=Path)
    parser.add_argument("--sampler-checkpoint", required=True)
    parser.add_argument("--manifest", type=Path, default=DEFAULT_CORPUS_MANIFEST)
    parser.add_argument("--split-manifest", type=Path, default=DEFAULT_SPLIT_MANIFEST)
    parser.add_argument("--out", required=True, type=Path)
    return parser.parse_args(argv)


def main(argv: Sequence[str] | None = None) -> int:
    args = parse_args(argv)
    try:
        try:
            validate_sampler_checkpoint(args.sampler_checkpoint)
            checkpoint_provenance = sampler_checkpoint_provenance(args.sampler_checkpoint)
        except ProviderConfigurationError as exc:
            raise ComparisonError(str(exc)) from None
        corpus = load_validated_corpus(args.manifest, args.split_manifest)
        validation_ids = [row.task_id for row in corpus.validation_rows]
        validation_prompt_hashes = {
            row.task_id: str(row.raw["provenance"]["prompt_sha256"])
            for row in corpus.validation_rows
        }
        comparison = compare_results(
            args.base,
            args.checkpoint,
            base_predictions_path=args.base_predictions,
            base_traces_path=args.base_traces,
            checkpoint_predictions_path=args.checkpoint_predictions,
            checkpoint_traces_path=args.checkpoint_traces,
            checkpoint_model_provenance=checkpoint_provenance,
            expected_task_ids=validation_ids,
            expected_prompt_sha256=validation_prompt_hashes,
        )
        comparison["experiment"] = {
            "base_model": MODEL_ID,
            "corpus_sha256": corpus.corpus_sha256,
            "dataset_manifest_sha256": corpus.dataset_manifest_sha256,
            "split_manifest_sha256": corpus.split_manifest_sha256,
            "sampler_checkpoint_sha256": hashlib.sha256(
                args.sampler_checkpoint.encode("utf-8")
            ).hexdigest(),
        }
        _write_new_json(args.out, comparison)
    except (ComparisonError, OSError, TrainingInputError) as exc:
        raise SystemExit(f"comparison error: {exc}") from exc
    print(json.dumps(comparison["checkpoint_minus_base"], sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
