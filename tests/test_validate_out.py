from __future__ import annotations

import json
import os
import tempfile
import unittest
from pathlib import Path

import openpyxl

from formulabench.artifacts import (
    atomic_copy,
    atomic_write_jsonl,
    load_dataset_manifest,
)
from formulabench.validate_out import validate_output

MODEL = "Qwen/Qwen3.8-27B"


def _write_workbook(path: Path, value: object) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    worksheet.title = "Sheet1"
    worksheet["A1"] = value
    workbook.save(path)
    workbook.close()


def _make_dataset(root: Path, task_ids: tuple[str, ...] = ("alpha-1", "beta-2")) -> Path:
    dataset = root / "dataset"
    records = []
    for index, task_id in enumerate(task_ids, start=1):
        relative_folder = f"spreadsheet/{task_id}"
        _write_workbook(
            dataset / relative_folder / f"1_{task_id}_init.xlsx",
            value=index,
        )
        records.append(
            {
                "id": task_id,
                "instruction": "Fill A1.",
                "spreadsheet_path": relative_folder,
                "instruction_type": "Cell-Level Manipulation",
                "answer_position": "A1",
                "answer_sheet": "Sheet1",
                "data_position": "A1",
            }
        )
    dataset.mkdir(parents=True, exist_ok=True)
    (dataset / "dataset.json").write_text(json.dumps(records), encoding="utf-8")
    return dataset


def _trace(*, model: str = MODEL, step: int = 1) -> dict:
    return {
        "step": step,
        "model": model,
        "prompt": "synthetic prompt",
        "response": '{"cells": []}',
        "input_tokens": 2,
        "output_tokens": 3,
        "latency_ms": 4,
        "error": None,
    }


def _make_valid_output(
    dataset: Path,
    out: Path,
    *,
    statuses: dict[str, str] | None = None,
) -> None:
    statuses = statuses or {}
    (out / "outputs").mkdir(parents=True)
    (out / "traces").mkdir()
    predictions = []
    for task in load_dataset_manifest(dataset):
        output = out / "outputs" / f"{task.id}.xlsx"
        atomic_copy(task.init_xlsx, output)
        status = statuses.get(task.id, "ok")
        predictions.append({"id": task.id, "output": f"outputs/{task.id}.xlsx", "status": status})
        trace = _trace()
        if status != "ok":
            trace["response"] = None
            trace["error"] = "synthetic failure"
        atomic_write_jsonl(out / "traces" / f"{task.id}.jsonl", [trace])
    atomic_write_jsonl(out / "predictions.jsonl", predictions)
    (out / "run.log").write_bytes(b"synthetic run\n")


def _codes(report: object) -> set[str]:
    return {issue.code for issue in report.issues}


class OutputValidationTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temporary = tempfile.TemporaryDirectory()
        self.root = Path(self.temporary.name)
        self.dataset = _make_dataset(self.root)
        self.out = self.root / "out"
        _make_valid_output(self.dataset, self.out)

    def tearDown(self) -> None:
        self.temporary.cleanup()

    def test_valid_output_passes_and_validation_is_read_only(self) -> None:
        paths = sorted(path for path in self.out.rglob("*") if path.is_file())
        before = {path.relative_to(self.out): path.read_bytes() for path in paths}

        report = validate_output(
            self.dataset,
            self.out,
            expected_model=MODEL,
        )

        after = {path.relative_to(self.out): path.read_bytes() for path in paths}
        self.assertTrue(report.ok, report.as_dict())
        self.assertEqual(report.expected_tasks, 2)
        self.assertEqual(report.prediction_records, 2)
        self.assertEqual(report.readable_workbooks, 2)
        self.assertEqual(report.valid_traces, 2)
        self.assertEqual(before, after)

    def test_exact_initial_workbook_is_valid_failure_fallback(self) -> None:
        replacement = self.root / "failure-out"
        _make_valid_output(
            self.dataset,
            replacement,
            statuses={"alpha-1": "error: synthetic"},
        )
        atomic_write_jsonl(replacement / "traces" / "alpha-1.jsonl", [])

        report = validate_output(
            self.dataset,
            replacement,
            expected_model=MODEL,
        )

        self.assertTrue(report.ok, report.as_dict())

    def test_ok_prediction_requires_a_successful_model_call(self) -> None:
        atomic_write_jsonl(self.out / "traces" / "alpha-1.jsonl", [])

        empty_report = validate_output(
            self.dataset,
            self.out,
            expected_model=MODEL,
        )

        self.assertIn("trace_no_successful_call", _codes(empty_report))

        failed_call = _trace()
        failed_call["response"] = None
        failed_call["error"] = "provider failure"
        atomic_write_jsonl(self.out / "traces" / "alpha-1.jsonl", [failed_call])
        failed_report = validate_output(
            self.dataset,
            self.out,
            expected_model=MODEL,
        )
        self.assertIn("trace_no_successful_call", _codes(failed_report))

    def test_trace_files_can_be_optional(self) -> None:
        for trace in (self.out / "traces").iterdir():
            trace.unlink()
        (self.out / "traces").rmdir()

        report = validate_output(
            self.dataset,
            self.out,
            expected_model=MODEL,
            require_traces=False,
        )

        self.assertTrue(report.ok, report.as_dict())

    def test_optional_existing_trace_is_still_validated(self) -> None:
        (self.out / "traces" / "beta-2.jsonl").unlink()
        atomic_write_jsonl(
            self.out / "traces" / "alpha-1.jsonl",
            [_trace(model="wrong-model")],
        )

        report = validate_output(
            self.dataset,
            self.out,
            expected_model=MODEL,
            require_traces=False,
        )

        self.assertIn("trace_model", _codes(report))
        self.assertNotIn("trace_missing", _codes(report))

    def test_modified_workbook_fails_failure_fallback_hash(self) -> None:
        predictions = [
            {
                "id": "alpha-1",
                "output": "outputs/alpha-1.xlsx",
                "status": "error: synthetic",
            },
            {
                "id": "beta-2",
                "output": "outputs/beta-2.xlsx",
                "status": "ok",
            },
        ]
        atomic_write_jsonl(self.out / "predictions.jsonl", predictions)
        workbook_path = self.out / "outputs" / "alpha-1.xlsx"
        workbook = openpyxl.load_workbook(workbook_path)
        workbook.active["A1"] = 999
        workbook.save(workbook_path)
        workbook.close()

        report = validate_output(
            self.dataset,
            self.out,
            expected_model=MODEL,
        )

        self.assertIn("failure_fallback_hash", _codes(report))

    def test_duplicate_missing_and_extra_predictions_are_all_reported(self) -> None:
        records = [
            {"id": "alpha-1", "output": "outputs/alpha-1.xlsx", "status": "ok"},
            {"id": "alpha-1", "output": "outputs/alpha-1.xlsx", "status": "ok"},
            {"id": "extra-3", "output": "outputs/extra-3.xlsx", "status": "ok"},
        ]
        atomic_write_jsonl(self.out / "predictions.jsonl", records)

        report = validate_output(
            self.dataset,
            self.out,
            expected_model=MODEL,
        )

        codes = _codes(report)
        self.assertIn("prediction_duplicate", codes)
        self.assertIn("prediction_missing", codes)
        self.assertIn("prediction_extra", codes)

    def test_prediction_traversal_is_rejected(self) -> None:
        outside = self.root / "outside.xlsx"
        _write_workbook(outside, 7)
        records = [
            {"id": "alpha-1", "output": "../outside.xlsx", "status": "ok"},
            {"id": "beta-2", "output": "outputs/beta-2.xlsx", "status": "ok"},
        ]
        atomic_write_jsonl(self.out / "predictions.jsonl", records)

        report = validate_output(
            self.dataset,
            self.out,
            expected_model=MODEL,
        )

        self.assertIn("workbook_missing_or_unsafe", _codes(report))
        self.assertIn("prediction_output_layout", _codes(report))

    @unittest.skipUnless(hasattr(os, "symlink"), "symbolic links are unavailable")
    def test_output_workbook_symlink_is_rejected(self) -> None:
        external = self.root / "external.xlsx"
        _write_workbook(external, 8)
        output = self.out / "outputs" / "alpha-1.xlsx"
        output.unlink()
        try:
            output.symlink_to(external)
        except OSError as exc:  # pragma: no cover - restricted Windows accounts.
            self.skipTest(str(exc))

        report = validate_output(
            self.dataset,
            self.out,
            expected_model=MODEL,
        )

        self.assertIn("workbook_missing_or_unsafe", _codes(report))

    def test_unreadable_workbook_is_rejected(self) -> None:
        (self.out / "outputs" / "alpha-1.xlsx").write_bytes(b"not a zip")

        report = validate_output(
            self.dataset,
            self.out,
            expected_model=MODEL,
        )

        self.assertIn("workbook_unreadable", _codes(report))

    def test_trace_schema_order_and_fixed_model_are_checked(self) -> None:
        invalid_trace = {
            "step": 2,
            "model": "another-model",
            "prompt": 5,
            "response": None,
            "input_tokens": -1,
            "output_tokens": None,
            "latency_ms": None,
            # error is intentionally absent.
        }
        atomic_write_jsonl(self.out / "traces" / "alpha-1.jsonl", [invalid_trace])

        report = validate_output(
            self.dataset,
            self.out,
            expected_model=MODEL,
        )

        codes = _codes(report)
        self.assertIn("trace_missing_fields", codes)
        self.assertIn("trace_step_order", codes)
        self.assertIn("trace_model", codes)
        self.assertIn("trace_field_type", codes)
        self.assertIn("trace_metric_type", codes)
        self.assertEqual(report.valid_traces, 1)

    def test_missing_and_extra_trace_files_are_rejected(self) -> None:
        (self.out / "traces" / "alpha-1.jsonl").unlink()
        atomic_write_jsonl(self.out / "traces" / "unknown.jsonl", [_trace()])

        report = validate_output(
            self.dataset,
            self.out,
            expected_model=MODEL,
        )

        self.assertIn("trace_missing", _codes(report))
        self.assertIn("trace_extra", _codes(report))

    def test_secret_scan_detects_value_without_echoing_it(self) -> None:
        dummy_secret = "dummy-secret-value-123456789"
        (self.out / "run.log").write_text(f"provider returned {dummy_secret}\n", encoding="utf-8")
        predictions = [
            {
                "id": "alpha-1",
                "output": f"outputs/{dummy_secret}.xlsx",
                "status": "ok",
            },
            {"id": "beta-2", "output": "outputs/beta-2.xlsx", "status": "ok"},
        ]
        atomic_write_jsonl(self.out / "predictions.jsonl", predictions)

        report = validate_output(
            self.dataset,
            self.out,
            expected_model=MODEL,
            secret_values=[dummy_secret],
        )

        serialised_report = json.dumps(report.as_dict(), sort_keys=True)
        self.assertIn("secret_leak", _codes(report))
        self.assertNotIn(dummy_secret, serialised_report)


if __name__ == "__main__":
    unittest.main()
