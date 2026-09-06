from __future__ import annotations

import hashlib
import json
from pathlib import Path

import openpyxl
import pytest

from formulabench.constants import MODEL_ID, MODEL_PROVENANCE
from formulabench.runner import _request_trace_metadata
from training.compare_results import (
    ComparisonError,
    _write_new_json,
    compare_results,
    parse_args,
)

CHECKPOINT = "tinker://01234567-89ab-cdef-0123-456789abcdef:train:0/sampler_weights/final"
CHECKPOINT_PROVENANCE = (
    "sampler_checkpoint_sha256:" + hashlib.sha256(CHECKPOINT.encode("utf-8")).hexdigest()
)
PROMPT = "synthetic prompt"
PROMPT_SHA256 = hashlib.sha256(PROMPT.encode("utf-8")).hexdigest()


def _summary(items: list[dict[str, object]]) -> dict[str, object]:
    graded = [item for item in items if item["status"] == "graded"]
    cell_total = sum(int(item["cells"]) for item in graded)

    def rate(rows: list[dict[str, object]]) -> float | None:
        return (
            round(sum(item.get("pass", False) is True for item in rows) / len(rows), 4)
            if rows
            else None
        )

    return {
        "items": len(items),
        "graded": len(graded),
        "missing": 0,
        "errors": sum(item["status"] == "error" for item in items),
        "pass_rate": rate(items),
        "cell_accuracy": (
            round(sum(int(item["correct"]) for item in graded) / cell_total, 4)
            if cell_total
            else None
        ),
        "pass_rate_cell_level": rate(
            [item for item in items if str(item["type"]).startswith("Cell")]
        ),
        "pass_rate_sheet_level": rate(
            [item for item in items if str(item["type"]).startswith("Sheet")]
        ),
    }


def _result(
    path: Path,
    ids: list[str],
    *,
    improved: bool,
    error_ids: set[str] | None = None,
) -> None:
    error_ids = error_ids or set()
    items: list[dict[str, object]] = []
    for index, task_id in enumerate(ids):
        if task_id in error_ids:
            items.append(
                {
                    "id": task_id,
                    "type": "Cell-Level Manipulation",
                    "status": "error",
                    "error": "synthetic evaluator error",
                }
            )
            continue
        correct = 2 if improved and index == 0 else 1
        items.append(
            {
                "id": task_id,
                "type": "Cell-Level Manipulation",
                "status": "graded",
                "cells": 2,
                "correct": correct,
                "pass": correct == 2,
                "mismatches": [] if correct == 2 else [{"cell": "Sheet1!A1"}],
            }
        )
    path.write_text(
        json.dumps(
            {
                "summary": _summary(items),
                "items": items,
            },
            sort_keys=True,
        ),
        encoding="utf-8",
    )


def _write_jsonl(path: Path, records: list[dict[str, object]]) -> None:
    path.write_text(
        "".join(json.dumps(record, sort_keys=True) + "\n" for record in records),
        encoding="utf-8",
    )


def _arm(
    root: Path,
    ids: list[str],
    provenance: str,
    *,
    statuses: dict[str, str] | None = None,
    empty_trace_ids: set[str] | None = None,
) -> tuple[Path, Path]:
    statuses = statuses or {}
    empty_trace_ids = empty_trace_ids or set()
    root.mkdir()
    outputs = root / "outputs"
    outputs.mkdir()
    traces = root / "traces"
    traces.mkdir()
    predictions: list[dict[str, object]] = []
    for task_id in ids:
        status = statuses.get(task_id, "ok")
        failure_codes = [] if status == "ok" else status.removeprefix("error:").split(",")
        predictions.append(
            {
                "id": task_id,
                "output": f"outputs/{task_id}.xlsx",
                "status": status,
                "failure_codes": failure_codes,
            }
        )
        workbook = openpyxl.Workbook()
        workbook.active["A1"] = task_id
        workbook.save(outputs / f"{task_id}.xlsx")
        workbook.close()
        request = _request_trace_metadata(provenance)
        if task_id in empty_trace_ids:
            _write_jsonl(traces / f"{task_id}.jsonl", [])
            continue
        _write_jsonl(
            traces / f"{task_id}.jsonl",
            [
                {
                    "step": 1,
                    "model": MODEL_ID,
                    "model_provenance": provenance,
                    "request": request,
                    "prompt": PROMPT,
                    "response": "{}" if status == "ok" else None,
                    "input_tokens": 3,
                    "output_tokens": 1,
                    "latency_ms": 5,
                    "error": None if status == "ok" else status,
                    "failure_codes": failure_codes,
                    "renderer": request["renderer"],
                    "stop_reason": "stop" if status == "ok" else None,
                    "parse_termination": "stop_sequence" if status == "ok" else None,
                    "response_format": "strict_json_fallback" if status == "ok" else None,
                    "response_rejection": None,
                    "context": (
                        {
                            "characters": 10,
                            "included_cells": 1,
                            "omitted_cells": 0,
                            "truncated": False,
                            "input_sha256": "c" * 64,
                        }
                        if status == "ok"
                        else None
                    ),
                }
            ],
        )
    predictions_path = root / "predictions.jsonl"
    _write_jsonl(predictions_path, predictions)
    return predictions_path, traces


def _comparison_fixture(
    tmp_path: Path,
    *,
    base_ids: list[str] | None = None,
    checkpoint_ids: list[str] | None = None,
    base_statuses: dict[str, str] | None = None,
    checkpoint_statuses: dict[str, str] | None = None,
    base_empty_trace_ids: set[str] | None = None,
    checkpoint_empty_trace_ids: set[str] | None = None,
    base_error_ids: set[str] | None = None,
    checkpoint_error_ids: set[str] | None = None,
) -> dict[str, Path]:
    base_ids = base_ids or ["a", "b"]
    checkpoint_ids = checkpoint_ids or ["a", "b"]
    base_result = tmp_path / "base-results.json"
    checkpoint_result = tmp_path / "checkpoint-results.json"
    _result(base_result, base_ids, improved=False, error_ids=base_error_ids)
    _result(
        checkpoint_result,
        checkpoint_ids,
        improved=True,
        error_ids=checkpoint_error_ids,
    )
    base_predictions, base_traces = _arm(
        tmp_path / "base",
        base_ids,
        MODEL_PROVENANCE,
        statuses=base_statuses,
        empty_trace_ids=base_empty_trace_ids,
    )
    checkpoint_predictions, checkpoint_traces = _arm(
        tmp_path / "checkpoint",
        checkpoint_ids,
        CHECKPOINT_PROVENANCE,
        statuses=checkpoint_statuses,
        empty_trace_ids=checkpoint_empty_trace_ids,
    )
    return {
        "base_result": base_result,
        "checkpoint_result": checkpoint_result,
        "base_predictions": base_predictions,
        "base_traces": base_traces,
        "checkpoint_predictions": checkpoint_predictions,
        "checkpoint_traces": checkpoint_traces,
    }


def _compare(paths: dict[str, Path], *, expected_task_ids: list[str] | None = None) -> dict:
    return compare_results(
        paths["base_result"],
        paths["checkpoint_result"],
        base_predictions_path=paths["base_predictions"],
        base_traces_path=paths["base_traces"],
        checkpoint_predictions_path=paths["checkpoint_predictions"],
        checkpoint_traces_path=paths["checkpoint_traces"],
        checkpoint_model_provenance=CHECKPOINT_PROVENANCE,
        expected_task_ids=expected_task_ids,
        expected_prompt_sha256={
            task_id: PROMPT_SHA256 for task_id in (expected_task_ids or ["a", "b"])
        },
    )


def test_comparison_is_task_aligned_provenance_and_hash_bound(
    tmp_path: Path,
) -> None:
    paths = _comparison_fixture(
        tmp_path,
        checkpoint_statuses={"b": "error:model_call_failed"},
    )

    result = _compare(paths, expected_task_ids=["a", "b"])

    assert result["schema_version"] == 2
    assert result["ordered_task_ids"] == ["a", "b"]
    assert result["task_count"] == 2
    assert result["evidence_validation"] == "provenance-and-hash-bound"
    assert result["checkpoint_minus_base"] == {
        "cell_accuracy": 0.25,
        "pass_rate": 0.5,
        "pass_rate_cell_level": 0.5,
        "pass_rate_sheet_level": None,
    }
    assert result["evidence"]["base"]["prediction_counts"] == {
        "accepted": 2,
        "empty_pre_call_failures": 0,
        "failed": 0,
        "total": 2,
    }
    assert result["evidence"]["checkpoint"]["prediction_counts"] == {
        "accepted": 1,
        "empty_pre_call_failures": 0,
        "failed": 1,
        "total": 2,
    }
    assert result["evidence"]["base"]["model_provenance"] == MODEL_PROVENANCE
    assert result["evidence"]["checkpoint"]["model_provenance"] == (CHECKPOINT_PROVENANCE)
    for key in (
        "predictions_sha256",
        "outputs_manifest_sha256",
        "traces_manifest_sha256",
        "prediction_trace_bindings_sha256",
    ):
        assert len(result["evidence"]["base"][key]) == 64
        assert len(result["evidence"]["checkpoint"][key]) == 64
    assert set(result["evidence"]["base"]["trace_files_sha256"]) == {"a", "b"}
    assert set(result["evidence"]["base"]["output_files_sha256"]) == {"a", "b"}
    assert (
        result["evidence"]["base"]["predictions_sha256"]
        == hashlib.sha256(paths["base_predictions"].read_bytes()).hexdigest()
    )
    assert result["evidence"]["checkpoint"]["trace_files_sha256"]["a"] == (
        hashlib.sha256((paths["checkpoint_traces"] / "a.jsonl").read_bytes()).hexdigest()
    )
    assert result["evaluator_errors"] == {
        "base": 0,
        "checkpoint": 0,
        "checkpoint_minus_base": 0,
        "counts_equal": True,
    }
    assert result["metric_comparability"]["cell_accuracy"] == {
        "comparable": True,
        "reason": None,
    }


def test_comparison_keeps_evaluator_only_api_compatibility(tmp_path: Path) -> None:
    base = tmp_path / "base.json"
    checkpoint = tmp_path / "checkpoint.json"
    _result(base, ["a", "b"], improved=False)
    _result(checkpoint, ["a", "b"], improved=True)

    result = compare_results(base, checkpoint)

    assert result["evidence_validation"] == "not-supplied"
    assert "evidence" not in result
    assert result["checkpoint_minus_base"]["cell_accuracy"] == 0.25


def test_comparison_rejects_bad_prediction_status(tmp_path: Path) -> None:
    paths = _comparison_fixture(tmp_path)
    predictions = paths["checkpoint_predictions"]
    records = [json.loads(line) for line in predictions.read_text().splitlines()]
    records[0]["status"] = "success"
    _write_jsonl(predictions, records)

    with pytest.raises(ComparisonError, match="invalid status"):
        _compare(paths)


@pytest.mark.parametrize("mutation", ["extra_field", "invented_failure", "duplicate_failure"])
def test_comparison_enforces_canonical_prediction_failure_schema(
    tmp_path: Path,
    mutation: str,
) -> None:
    paths = _comparison_fixture(tmp_path)
    predictions = paths["checkpoint_predictions"]
    records = [json.loads(line) for line in predictions.read_text().splitlines()]
    if mutation == "extra_field":
        records[0]["model"] = MODEL_ID
    elif mutation == "invented_failure":
        records[0]["status"] = "error:invented_failure"
        records[0]["failure_codes"] = ["invented_failure"]
    else:
        records[0]["status"] = "error:model_call_failed,model_call_failed"
        records[0]["failure_codes"] = ["model_call_failed", "model_call_failed"]
    _write_jsonl(predictions, records)

    with pytest.raises(ComparisonError, match="invalid schema|invalid failure codes"):
        _compare(paths)


def test_comparison_rejects_zero_accepted_arm(tmp_path: Path) -> None:
    paths = _comparison_fixture(
        tmp_path,
        checkpoint_statuses={
            "a": "error:model_call_failed",
            "b": "error:invalid_response_schema",
        },
    )

    with pytest.raises(ComparisonError, match="checkpoint arm has zero accepted"):
        _compare(paths)


def test_comparison_accepts_and_records_mixed_pre_call_failure(tmp_path: Path) -> None:
    paths = _comparison_fixture(
        tmp_path,
        base_statuses={"b": "error:context_build_failed"},
        checkpoint_statuses={"b": "error:context_build_failed"},
        base_empty_trace_ids={"b"},
        checkpoint_empty_trace_ids={"b"},
    )

    result = _compare(paths)

    assert result["evidence"]["base"]["prediction_counts"] == {
        "accepted": 1,
        "empty_pre_call_failures": 1,
        "failed": 1,
        "total": 2,
    }
    assert result["evidence"]["base"]["trace_files_sha256"]["b"] == hashlib.sha256(b"").hexdigest()


def test_comparison_rejects_empty_trace_for_non_context_failure(tmp_path: Path) -> None:
    paths = _comparison_fixture(
        tmp_path,
        base_statuses={"b": "error:model_call_failed"},
        base_empty_trace_ids={"b"},
    )

    with pytest.raises(ComparisonError, match="empty without a pre-call context failure"):
        _compare(paths)


def test_comparison_rejects_model_call_trace_for_context_failure(tmp_path: Path) -> None:
    paths = _comparison_fixture(
        tmp_path,
        base_statuses={"b": "error:context_build_failed"},
    )

    with pytest.raises(ComparisonError, match="model call after a context failure"):
        _compare(paths)


def test_comparison_rejects_trace_provenance_mismatch(tmp_path: Path) -> None:
    paths = _comparison_fixture(tmp_path)
    trace_path = paths["checkpoint_traces"] / "a.jsonl"
    record = json.loads(trace_path.read_text())
    record["model_provenance"] = MODEL_PROVENANCE
    _write_jsonl(trace_path, [record])

    with pytest.raises(ComparisonError, match="wrong model provenance"):
        _compare(paths)


def test_comparison_rejects_request_provenance_mismatch(tmp_path: Path) -> None:
    paths = _comparison_fixture(tmp_path)
    trace_path = paths["base_traces"] / "a.jsonl"
    record = json.loads(trace_path.read_text())
    record["request"]["model_provenance"] = CHECKPOINT_PROVENANCE
    _write_jsonl(trace_path, [record])

    with pytest.raises(ComparisonError, match="wrong fixed request settings"):
        _compare(paths)


def test_comparison_rejects_changed_fixed_request_settings(tmp_path: Path) -> None:
    paths = _comparison_fixture(tmp_path)
    trace_path = paths["checkpoint_traces"] / "a.jsonl"
    record = json.loads(trace_path.read_text())
    record["request"]["temperature"] = 0.5
    _write_jsonl(trace_path, [record])

    with pytest.raises(ComparisonError, match="wrong fixed request settings"):
        _compare(paths)


def test_comparison_rejects_different_or_non_corpus_prompts(tmp_path: Path) -> None:
    paths = _comparison_fixture(tmp_path)
    trace_path = paths["checkpoint_traces"] / "a.jsonl"
    record = json.loads(trace_path.read_text())
    record["prompt"] = "different prompt"
    _write_jsonl(trace_path, [record])

    with pytest.raises(ComparisonError, match="validated corpus prompt"):
        _compare(paths)


def test_comparison_rejects_different_workbook_context(tmp_path: Path) -> None:
    paths = _comparison_fixture(tmp_path)
    base_trace = paths["base_traces"] / "a.jsonl"
    checkpoint_trace = paths["checkpoint_traces"] / "a.jsonl"
    base_record = json.loads(base_trace.read_text())
    checkpoint_record = json.loads(checkpoint_trace.read_text())
    base_record["context"]["input_sha256"] = "a" * 64
    checkpoint_record["context"]["input_sha256"] = "b" * 64
    _write_jsonl(base_trace, [base_record])
    _write_jsonl(checkpoint_trace, [checkpoint_record])

    with pytest.raises(ComparisonError, match="different workbook context"):
        _compare(paths)


def test_comparison_rejects_null_context_input_hash(tmp_path: Path) -> None:
    paths = _comparison_fixture(tmp_path)
    trace_path = paths["checkpoint_traces"] / "a.jsonl"
    record = json.loads(trace_path.read_text())
    record["context"]["input_sha256"] = None
    _write_jsonl(trace_path, [record])

    with pytest.raises(ComparisonError, match="invalid context provenance"):
        _compare(paths)


def test_comparison_rejects_more_than_one_model_call_per_task(tmp_path: Path) -> None:
    paths = _comparison_fixture(tmp_path)
    trace_path = paths["base_traces"] / "a.jsonl"
    first = json.loads(trace_path.read_text())
    second = dict(first)
    second["step"] = 2
    _write_jsonl(trace_path, [first, second])

    with pytest.raises(ComparisonError, match="exactly one model call"):
        _compare(paths)


def test_comparison_rejects_prediction_task_mismatch_or_reordering(tmp_path: Path) -> None:
    paths = _comparison_fixture(tmp_path)
    predictions = paths["checkpoint_predictions"]
    records = [json.loads(line) for line in predictions.read_text().splitlines()]
    _write_jsonl(predictions, list(reversed(records)))

    with pytest.raises(ComparisonError, match="exact ordered comparison tasks"):
        _compare(paths)


def test_output_workbook_bytes_are_bound_and_missing_outputs_are_rejected(
    tmp_path: Path,
) -> None:
    paths = _comparison_fixture(tmp_path)
    initial = _compare(paths)
    output_path = paths["checkpoint_predictions"].parent / "outputs" / "a.xlsx"
    workbook = openpyxl.load_workbook(output_path)
    workbook.active["A1"] = "changed after evaluation"
    workbook.save(output_path)
    workbook.close()

    changed = _compare(paths)

    assert (
        initial["evidence"]["checkpoint"]["output_files_sha256"]["a"]
        != changed["evidence"]["checkpoint"]["output_files_sha256"]["a"]
    )
    assert (
        initial["evidence"]["checkpoint"]["prediction_trace_bindings_sha256"]
        != changed["evidence"]["checkpoint"]["prediction_trace_bindings_sha256"]
    )

    output_path.unlink()
    with pytest.raises(ComparisonError, match="outputs do not exactly match"):
        _compare(paths)


def test_comparison_rejects_different_or_reordered_evaluator_tasks(tmp_path: Path) -> None:
    paths = _comparison_fixture(tmp_path, checkpoint_ids=["b", "a"])

    with pytest.raises(ComparisonError, match="same ordered tasks"):
        _compare(paths)


@pytest.mark.parametrize("mutation", ["type", "cells"])
def test_comparison_rejects_different_evaluator_populations_across_arms(
    tmp_path: Path,
    mutation: str,
) -> None:
    paths = _comparison_fixture(tmp_path)
    payload = json.loads(paths["checkpoint_result"].read_text())
    item = payload["items"][1]
    if mutation == "type":
        item["type"] = "Sheet-Level Manipulation"
        expected = "different evaluator types"
    else:
        item["cells"] = 3
        expected = "different graded-cell counts"
    payload["summary"] = _summary(payload["items"])
    paths["checkpoint_result"].write_text(json.dumps(payload), encoding="utf-8")

    with pytest.raises(ComparisonError, match=expected):
        _compare(paths)


def test_comparison_rejects_tasks_outside_validated_partition(tmp_path: Path) -> None:
    paths = _comparison_fixture(tmp_path)

    with pytest.raises(ComparisonError, match="validated corpus partition"):
        _compare(paths, expected_task_ids=["a", "c"])


@pytest.mark.parametrize(
    ("base_errors", "checkpoint_errors", "error_delta", "counts_equal"),
    [
        ({"a"}, set(), -1, False),
        ({"a"}, {"b"}, 0, True),
    ],
)
def test_evaluator_errors_make_cell_accuracy_delta_non_comparable(
    tmp_path: Path,
    base_errors: set[str],
    checkpoint_errors: set[str],
    error_delta: int,
    counts_equal: bool,
) -> None:
    paths = _comparison_fixture(
        tmp_path,
        base_error_ids=base_errors,
        checkpoint_error_ids=checkpoint_errors,
    )

    result = _compare(paths)
    base_summary = json.loads(paths["base_result"].read_text())["summary"]
    checkpoint_summary = json.loads(paths["checkpoint_result"].read_text())["summary"]

    assert result["base"] == base_summary
    assert result["checkpoint"] == checkpoint_summary
    assert result["checkpoint_minus_base"]["cell_accuracy"] is None
    assert result["evaluator_errors"] == {
        "base": len(base_errors),
        "checkpoint": len(checkpoint_errors),
        "checkpoint_minus_base": error_delta,
        "counts_equal": counts_equal,
    }
    comparability = result["metric_comparability"]["cell_accuracy"]
    assert comparability["comparable"] is False
    assert "excludes evaluator-error tasks" in comparability["reason"]


def test_comparison_rejects_inconsistent_evaluator_error_count(tmp_path: Path) -> None:
    paths = _comparison_fixture(tmp_path, checkpoint_error_ids={"a"})
    payload = json.loads(paths["checkpoint_result"].read_text())
    payload["summary"]["errors"] = 0
    paths["checkpoint_result"].write_text(json.dumps(payload), encoding="utf-8")

    with pytest.raises(ComparisonError, match="invalid errors"):
        _compare(paths)


def test_comparison_recomputes_and_rejects_evaluator_summary_arithmetic(
    tmp_path: Path,
) -> None:
    paths = _comparison_fixture(tmp_path)
    payload = json.loads(paths["checkpoint_result"].read_text())
    payload["summary"]["cell_accuracy"] = 1.0
    paths["checkpoint_result"].write_text(json.dumps(payload), encoding="utf-8")

    with pytest.raises(ComparisonError, match="invalid cell_accuracy"):
        _compare(paths)


def test_comparison_cli_requires_explicit_prediction_and_trace_evidence() -> None:
    common = [
        "--base=base.json",
        "--checkpoint=checkpoint.json",
        f"--sampler-checkpoint={CHECKPOINT}",
        "--out=comparison.json",
    ]
    with pytest.raises(SystemExit):
        parse_args(common)

    args = parse_args(
        [
            *common,
            "--base-predictions=base/predictions.jsonl",
            "--base-traces=base/traces",
            "--checkpoint-predictions=checkpoint/predictions.jsonl",
            "--checkpoint-traces=checkpoint/traces",
        ]
    )
    assert args.base_predictions == Path("base/predictions.jsonl")
    assert args.base_traces == Path("base/traces")
    assert args.checkpoint_predictions == Path("checkpoint/predictions.jsonl")
    assert args.checkpoint_traces == Path("checkpoint/traces")


def test_comparison_output_never_overwrites_existing_evidence(tmp_path: Path) -> None:
    output = tmp_path / "comparison.json"
    output.write_text("retained\n", encoding="utf-8")

    with pytest.raises(ComparisonError, match="new path"):
        _write_new_json(output, {"schema_version": 2})
    assert output.read_text(encoding="utf-8") == "retained\n"
