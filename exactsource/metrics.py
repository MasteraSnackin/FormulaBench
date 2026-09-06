"""Deterministic, trace-derived benchmark run metrics."""

from __future__ import annotations

import json
import math
from collections import Counter
from collections.abc import Mapping, Sequence
from dataclasses import dataclass, field
from typing import Any

from exactsource.artifacts import (
    ArtifactError,
    OutputLayout,
    Prediction,
    atomic_write_text,
    read_jsonl,
)

METRICS_SCHEMA_VERSION = 2
_MEASUREMENTS = (
    "input_tokens",
    "output_tokens",
    "cache_creation_input_tokens",
    "cache_read_input_tokens",
    "provider_latency_ms",
)
_TRACE_FIELD = {
    "input_tokens": "input_tokens",
    "output_tokens": "output_tokens",
    "cache_creation_input_tokens": "cache_creation_input_tokens",
    "cache_read_input_tokens": "cache_read_input_tokens",
    "provider_latency_ms": "latency_ms",
}
_CHARACTER_FIELDS = (
    "message_chars",
    "request_chars",
    "response_chars",
    "answer_chars",
    "thinking_chars",
)
_STAGE_FIELDS = (
    "context_build_latency_ms",
    "message_build_latency_ms",
    "plan_parse_latency_ms",
    "plan_apply_latency_ms",
)


@dataclass(slots=True)
class _KnownUnknownTotal:
    known_sum: int = 0
    known_attempts: int = 0
    unknown_attempts: int = 0

    def observe(self, record: Mapping[str, Any], field: str, *, location: str) -> None:
        value = record.get(field)
        if value is None:
            self.unknown_attempts += 1
            return
        _require_non_negative_integer(value, field=field, location=location)
        self.known_sum += value
        self.known_attempts += 1

    def as_dict(self) -> dict[str, int]:
        return {
            "known_sum": self.known_sum,
            "known_attempts": self.known_attempts,
            "unknown_attempts": self.unknown_attempts,
        }


@dataclass(slots=True)
class _Distribution:
    """Integer observations with exact median and nearest-rank p95."""

    values: list[int] = field(default_factory=list)
    unknown_count: int = 0

    def observe(self, record: Mapping[str, Any], field: str, *, location: str) -> None:
        self.observe_value(record.get(field), field=field, location=location)

    def observe_value(self, value: object, *, field: str, location: str) -> None:
        if value is None:
            self.unknown_count += 1
            return
        _require_non_negative_integer(value, field=field, location=location)
        self.values.append(value)

    def as_dict(self) -> dict[str, int | float | None]:
        ordered = sorted(self.values)
        if not ordered:
            minimum: int | None = None
            median: int | float | None = None
            p95: int | None = None
            maximum: int | None = None
        else:
            minimum = ordered[0]
            maximum = ordered[-1]
            middle = len(ordered) // 2
            if len(ordered) % 2:
                median = ordered[middle]
            else:
                median = (ordered[middle - 1] + ordered[middle]) / 2
            p95 = ordered[math.ceil(0.95 * len(ordered)) - 1]
        return {
            "known_count": len(ordered),
            "unknown_count": self.unknown_count,
            "sum": sum(ordered),
            "min": minimum,
            "median": median,
            "p95": p95,
            "max": maximum,
        }


def build_run_metrics(
    layout: OutputLayout,
    predictions: Sequence[Prediction],
    *,
    run_wall_time_ms: int,
    task_latencies_ms: Mapping[str, int | None] | None = None,
) -> dict[str, Any]:
    """Aggregate final trace files without sharing counters across workers.

    A trace record represents one provider attempt. Logical model calls are grouped
    by task and ``semantic_attempt``; a legacy or injected trace without that field
    is conservatively treated as its own call. ``input_tokens`` is summed exactly as
    recorded because the Tinker adapter already includes cached input in that value.
    Cache fields are reported separately and are never added to the input total.

    Distribution percentiles are deterministic: median is exact (and may be a
    half-integer), while p95 uses the nearest-rank definition.
    """

    _require_non_negative_integer(
        run_wall_time_ms,
        field="run wall time",
        location="metrics input",
    )
    prediction_ids = [prediction.id for prediction in predictions]
    if len(set(prediction_ids)) != len(prediction_ids):
        raise ArtifactError("metrics predictions contain duplicate task ids")
    if task_latencies_ms is not None:
        expected = set(prediction_ids)
        actual = set(task_latencies_ms)
        if actual != expected:
            raise ArtifactError("task latency ids do not match prediction ids")

    totals = {name: _KnownUnknownTotal() for name in _MEASUREMENTS}
    characters = {name: _Distribution() for name in _CHARACTER_FIELDS}
    stages = {name: _Distribution() for name in _STAGE_FIELDS}
    statuses: Counter[str] = Counter()
    provider_statuses: Counter[str] = Counter()
    plan_statuses: Counter[str] = Counter()
    call_records: dict[
        tuple[str, str, int],
        list[tuple[Mapping[str, Any], str]],
    ] = {}
    task_records: dict[str, list[tuple[Mapping[str, Any], str]]] = {}
    provider_attempt_latency = _Distribution()
    provider_latency_by_status: dict[str, _Distribution] = {}
    retry_delay = _Distribution()
    transport_retries = 0
    attempts = 0

    for prediction in predictions:
        records = read_jsonl(layout.trace_path(prediction.id))
        task_records[prediction.id] = []
        for record_index, record in enumerate(records, start=1):
            attempts += 1
            location = f"trace {prediction.id!r} record {record_index}"
            task_records[prediction.id].append((record, location))

            status = _status(record, location=location)
            statuses[status] += 1
            provider_status = _provider_status(record, location=location)
            provider_statuses[provider_status] += 1

            semantic_attempt = record.get("semantic_attempt")
            if semantic_attempt is None:
                call_key = (prediction.id, "record", record_index)
            else:
                _require_positive_integer(
                    semantic_attempt,
                    field="semantic_attempt",
                    location=location,
                )
                call_key = (prediction.id, "semantic", semantic_attempt)
            call_records.setdefault(call_key, []).append((record, location))

            for metric_name, total in totals.items():
                total.observe(record, _TRACE_FIELD[metric_name], location=location)
            for field_name, distribution in characters.items():
                distribution.observe(record, field_name, location=location)

            provider_attempt_latency.observe(record, "latency_ms", location=location)
            provider_latency_by_status.setdefault(provider_status, _Distribution()).observe(
                record,
                "latency_ms",
                location=location,
            )
            if provider_status == "retry":
                transport_retries += 1
                retry_delay.observe_value(
                    _transport_retry_delay_ms(record, location=location),
                    field="transport_retry_delay_ms",
                    location=location,
                )

    task_latency = _Distribution()
    task_latency_by_outcome = {
        "succeeded": _Distribution(),
        "failed": _Distribution(),
    }
    task_latency_by_id: dict[str, int | None] = {}
    for prediction in predictions:
        records = task_records[prediction.id]
        traced_value = records[-1][0].get("task_latency_ms") if records else None
        value: int | None
        if task_latencies_ms is not None:
            value = task_latencies_ms[prediction.id]
            if traced_value is not None and value != traced_value:
                raise ArtifactError(
                    f"task {prediction.id!r} latency does not match its terminal trace"
                )
        else:
            value = traced_value
        task_latency_by_id[prediction.id] = value
        outcome = "succeeded" if prediction.status == "ok" else "failed"
        location = f"task {prediction.id!r}"
        task_latency.observe_value(value, field="task_latency_ms", location=location)
        if value is not None and value > run_wall_time_ms:
            raise ArtifactError(f"{location} latency exceeds the run wall time")
        task_latency_by_outcome[outcome].observe_value(
            value,
            field="task_latency_ms",
            location=location,
        )

        first_record = records[0] if records else None
        for field_name in ("context_build_latency_ms", "message_build_latency_ms"):
            value = _first_present(records, field_name)
            stage_location = first_record[1] if first_record is not None else location
            stages[field_name].observe_value(
                value,
                field=field_name,
                location=stage_location,
            )

    logical_call_latency = _Distribution()
    logical_latency_by_plan_status: dict[str, _Distribution] = {}
    semantic_repairs = 0
    semantic_repair_counts: Counter[str] = Counter()
    unknown_semantic_attempt_calls = 0
    for call_key, records in call_records.items():
        terminal, location = records[-1]
        plan_status = _plan_status(terminal, location=location)
        plan_statuses[plan_status] += 1

        logical_call_latency.observe(
            terminal,
            "logical_call_latency_ms",
            location=location,
        )
        logical_latency_by_plan_status.setdefault(plan_status, _Distribution()).observe(
            terminal,
            "logical_call_latency_ms",
            location=location,
        )

        if call_key[1] == "semantic":
            if call_key[2] > 1:
                semantic_repairs += 1
                recovery_reason = terminal.get("recovery_reason")
                if recovery_reason == "max_tokens":
                    semantic_repair_counts["max_tokens_recovery"] += 1
                elif recovery_reason is None and terminal.get("semantic_repair") is True:
                    semantic_repair_counts["ordinary"] += 1
                else:
                    semantic_repair_counts["other_or_unknown"] += 1
        else:
            unknown_semantic_attempt_calls += 1

        if plan_status != "not_reached":
            stages["plan_parse_latency_ms"].observe(
                terminal,
                "plan_parse_latency_ms",
                location=location,
            )
        if plan_status in {"accepted", "apply_rejected"}:
            stages["plan_apply_latency_ms"].observe(
                terminal,
                "plan_apply_latency_ms",
                location=location,
            )

    succeeded = sum(prediction.status == "ok" for prediction in predictions)
    return {
        "schema_version": METRICS_SCHEMA_VERSION,
        "timing_evidence": {
            "run_wall_time_ms": {
                "source": "coordinator monotonic clock",
                "trace_recomputable": False,
            },
            "task_latency_ms": {
                "source": "coordinator monotonic clock",
                "terminal_trace_cross_check": True,
            },
        },
        "run_wall_time_ms": run_wall_time_ms,
        "tasks": {
            "total": len(predictions),
            "succeeded": succeeded,
            "failed": len(predictions) - succeeded,
        },
        "model": {
            "calls": len(call_records),
            "attempts": attempts,
            "attempt_status_counts": dict(sorted(statuses.items())),
        },
        "usage": {name: totals[name].as_dict() for name in _MEASUREMENTS},
        "latency_ms": {
            "task": {
                "all": task_latency.as_dict(),
                "by_task": dict(sorted(task_latency_by_id.items())),
                "by_outcome": {
                    name: distribution.as_dict()
                    for name, distribution in sorted(task_latency_by_outcome.items())
                },
            },
            "logical_call": {
                "all": logical_call_latency.as_dict(),
                "by_plan_status": {
                    name: distribution.as_dict()
                    for name, distribution in sorted(logical_latency_by_plan_status.items())
                },
            },
            "provider_attempt": {
                "all": provider_attempt_latency.as_dict(),
                "by_provider_status": {
                    name: distribution.as_dict()
                    for name, distribution in sorted(provider_latency_by_status.items())
                },
            },
            "stages": {name: stages[name].as_dict() for name in _STAGE_FIELDS},
            "transport_retry_delay": retry_delay.as_dict(),
        },
        "characters": {
            "scope": "provider_attempt",
            **{name: characters[name].as_dict() for name in _CHARACTER_FIELDS},
        },
        "reliability": {
            "provider_status_counts": dict(sorted(provider_statuses.items())),
            "plan_status_counts": dict(sorted(plan_statuses.items())),
            "semantic_repairs": semantic_repairs,
            "semantic_repair_counts": {
                name: semantic_repair_counts[name]
                for name in (
                    "ordinary",
                    "max_tokens_recovery",
                    "other_or_unknown",
                )
            },
            "semantic_attempt_unknown_calls": unknown_semantic_attempt_calls,
            "transport_retries": transport_retries,
        },
    }


def write_run_metrics(
    layout: OutputLayout,
    predictions: Sequence[Prediction],
    *,
    run_wall_time_ms: int,
    task_latencies_ms: Mapping[str, int | None] | None = None,
) -> dict[str, Any]:
    """Build and atomically publish the required root metrics artefact."""

    metrics = build_run_metrics(
        layout,
        predictions,
        run_wall_time_ms=run_wall_time_ms,
        task_latencies_ms=task_latencies_ms,
    )
    text = json.dumps(metrics, ensure_ascii=False, indent=2, sort_keys=True) + "\n"
    atomic_write_text(layout.run_metrics_path, text)
    return metrics


def _require_non_negative_integer(value: object, *, field: str, location: str) -> None:
    if isinstance(value, bool) or not isinstance(value, int) or value < 0:
        raise ArtifactError(f"{location} has invalid {field}")


def _require_positive_integer(value: object, *, field: str, location: str) -> None:
    _require_non_negative_integer(value, field=field, location=location)
    if value < 1:
        raise ArtifactError(f"{location} has invalid {field}")


def _status(record: Mapping[str, Any], *, location: str) -> str:
    value = record.get("status")
    if value is None:
        return "unknown"
    if not isinstance(value, str) or not value:
        raise ArtifactError(f"{location} has invalid status")
    return value


def _provider_status(record: Mapping[str, Any], *, location: str) -> str:
    value = record.get("provider_status")
    if value is not None:
        if not isinstance(value, str) or not value:
            raise ArtifactError(f"{location} has invalid provider_status")
        return value
    legacy_status = _status(record, location=location)
    if legacy_status == "plan_rejected":
        return "success"
    if legacy_status in {"success", "retry", "error"}:
        return legacy_status
    return "unknown"


def _plan_status(record: Mapping[str, Any], *, location: str) -> str:
    value = record.get("plan_status")
    if value is not None:
        if not isinstance(value, str) or not value:
            raise ArtifactError(f"{location} has invalid plan_status")
        return value
    status = _status(record, location=location)
    if status == "plan_rejected":
        return "rejected"
    if (
        record.get("event") == "model_call"
        and _provider_status(record, location=location) == "success"
        and record.get("error") is None
    ):
        return "accepted"
    return "not_reached"


def _transport_retry_delay_ms(
    record: Mapping[str, Any],
    *,
    location: str,
) -> int | None:
    value = record.get("transport_retry_delay_ms")
    if value is not None:
        _require_non_negative_integer(
            value,
            field="transport_retry_delay_ms",
            location=location,
        )
        return value
    seconds = record.get("retry_delay_seconds")
    if seconds is None:
        return None
    if isinstance(seconds, bool) or not isinstance(seconds, (int, float)):
        raise ArtifactError(f"{location} has invalid retry_delay_seconds")
    if not math.isfinite(seconds) or seconds < 0:
        raise ArtifactError(f"{location} has invalid retry_delay_seconds")
    return max(0, round(seconds * 1_000))


def _first_present(
    records: Sequence[tuple[Mapping[str, Any], str]],
    field: str,
) -> object:
    for record, _ in records:
        if record.get(field) is not None:
            return record[field]
    return None
