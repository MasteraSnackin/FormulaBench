from __future__ import annotations

import json
from pathlib import Path

import pytest

from training.compare_results import ComparisonError, _write_new_json, compare_results


def _result(path: Path, ids: list[str], pass_rate: float, cell_accuracy: float) -> None:
    path.write_text(
        json.dumps(
            {
                "summary": {
                    "items": len(ids),
                    "graded": len(ids),
                    "pass_rate": pass_rate,
                    "cell_accuracy": cell_accuracy,
                    "pass_rate_cell_level": pass_rate,
                    "pass_rate_sheet_level": None,
                },
                "items": [{"id": task_id, "status": "graded"} for task_id in ids],
            },
            sort_keys=True,
        ),
        encoding="utf-8",
    )


def test_comparison_is_task_aligned_hash_bound_and_reports_checkpoint_delta(
    tmp_path: Path,
) -> None:
    base = tmp_path / "base.json"
    checkpoint = tmp_path / "checkpoint.json"
    _result(base, ["a", "b"], 0.25, 0.5)
    _result(checkpoint, ["a", "b"], 0.5, 0.75)

    result = compare_results(base, checkpoint)

    assert result["ordered_task_ids"] == ["a", "b"]
    assert result["task_count"] == 2
    assert len(result["base_results_sha256"]) == 64
    assert len(result["checkpoint_results_sha256"]) == 64
    assert result["checkpoint_minus_base"] == {
        "cell_accuracy": 0.25,
        "pass_rate": 0.25,
        "pass_rate_cell_level": 0.25,
        "pass_rate_sheet_level": None,
    }


def test_comparison_rejects_different_or_reordered_tasks(tmp_path: Path) -> None:
    base = tmp_path / "base.json"
    checkpoint = tmp_path / "checkpoint.json"
    _result(base, ["a", "b"], 0.25, 0.5)
    _result(checkpoint, ["b", "a"], 0.5, 0.75)

    with pytest.raises(ComparisonError, match="same ordered tasks"):
        compare_results(base, checkpoint)


def test_comparison_rejects_tasks_outside_validated_partition(tmp_path: Path) -> None:
    base = tmp_path / "base.json"
    checkpoint = tmp_path / "checkpoint.json"
    _result(base, ["a", "b"], 0.25, 0.5)
    _result(checkpoint, ["a", "b"], 0.5, 0.75)

    with pytest.raises(ComparisonError, match="validated corpus partition"):
        compare_results(base, checkpoint, expected_task_ids=["a", "c"])


def test_comparison_output_never_overwrites_existing_evidence(tmp_path: Path) -> None:
    output = tmp_path / "comparison.json"
    output.write_text("retained\n", encoding="utf-8")

    with pytest.raises(ComparisonError, match="new path"):
        _write_new_json(output, {"schema_version": 1})
    assert output.read_text(encoding="utf-8") == "retained\n"
