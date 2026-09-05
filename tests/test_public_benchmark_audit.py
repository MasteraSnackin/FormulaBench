from __future__ import annotations

import datetime as dt
import json
import os
from collections.abc import Iterator
from pathlib import Path

import openpyxl
import pytest
from openpyxl.utils.datetime import CALENDAR_MAC_1904

from experiments.analyse_public_benchmark import analyse, canonical_json


def _save_workbook(path: Path, *, task: str, golden: bool) -> None:
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    if task == "alpha":
        worksheet.title = "Secret ledger"
        worksheet["A1"] = "UNCHANGED_ANSWER_SECRET"
        worksheet["B1"] = None if golden else "CLEAR_THIS_ANSWER_SECRET"
        worksheet["C1"] = dt.datetime(2024, 1, 2)
        worksheet["D1"] = "01/02/2024"
        worksheet["E1"] = "OUTSIDE_TARGET_SECRET"
    else:
        workbook.epoch = CALENDAR_MAC_1904
        worksheet.title = "Secret output"
        worksheet["A1"] = dt.datetime(2024, 3, 4, 12, 30)
        worksheet["B1"] = dt.time(9, 15)
        worksheet[f"C{3 if golden else 2}"] = "PRIVATE_DYNAMIC_EXTENT"
    workbook.save(path)
    workbook.close()


def _synthetic_dataset(tmp_path: Path) -> Path:
    tasks = [
        {
            "answer_position": "A1:B1",
            "answer_sheet": "Secret ledger",
            "id": "private-alpha",
            "instruction": "Do not publish this private alpha instruction.",
            "instruction_type": "Cell-Level Manipulation",
            "spreadsheet_path": "spreadsheet/private-alpha",
        },
        {
            "answer_position": "A:B",
            "answer_sheet": "Secret output",
            "id": "private-beta",
            "instruction": "Do not publish this private beta instruction.",
            "instruction_type": "Sheet-Level Manipulation",
            "spreadsheet_path": "spreadsheet/private-beta",
        },
    ]
    (tmp_path / "dataset.json").write_text(json.dumps(tasks), encoding="utf-8")
    for task_name in ("alpha", "beta"):
        folder = tmp_path / "spreadsheet" / f"private-{task_name}"
        folder.mkdir(parents=True)
        _save_workbook(folder / f"1_private-{task_name}_init.xlsx", task=task_name, golden=False)
        _save_workbook(
            folder / f"1_private-{task_name}_golden.xlsx",
            task=task_name,
            golden=True,
        )
    return tmp_path


def _lists_in(value: object) -> Iterator[list[object]]:
    if isinstance(value, dict):
        for child in value.values():
            yield from _lists_in(child)
    elif isinstance(value, list):
        yield value
        for child in value:
            yield from _lists_in(child)


def test_synthetic_audit_matches_evaluator_semantics_and_emits_only_aggregates(
    tmp_path: Path,
) -> None:
    payload = analyse(_synthetic_dataset(tmp_path))
    rendered = canonical_json(payload)

    assert payload["scope"] == {
        "instruction_levels": {
            "cell_level": {
                "instruction_type": "Cell-Level Manipulation",
                "tasks": 1,
            },
            "sheet_level": {
                "instruction_type": "Sheet-Level Manipulation",
                "tasks": 1,
            },
        },
        "tasks": 2,
    }
    assert payload["targets"] | {"by_instruction_level": None} == {
        "by_instruction_level": None,
        "dynamic_cells": 6,
        "dynamic_tasks": 1,
        "finite_cells": 2,
        "finite_size_buckets": {
            "1-100": 1,
            "101-500": 0,
            "501-1000": 0,
            "1001-10000": 0,
            "10001+": 0,
        },
        "finite_tasks": 1,
        "largest_finite_task_cells": 2,
        "scored_cells_including_dynamic": 8,
    }
    assert payload["golden_targets"] | {"by_instruction_level": None} == {
        "blank_cell_fraction": 0.625,
        "blank_cells": 5,
        "by_instruction_level": None,
        "cells": 8,
        "tasks": 2,
        "tasks_with_blank_cells": 2,
    }
    assert payload["initial_target_state"] | {"by_instruction_level": None} == {
        "by_instruction_level": None,
        "correct_nonblank_prefill_cells": 3,
        "initial_nonblank_target_cells": 4,
        "nonblank_cells_requiring_clearing": 1,
        "tasks_requiring_nonblank_cell_clearing": 1,
        "tasks_with_correct_nonblank_prefill_and_remaining_changes": 1,
        "tasks_with_initial_nonblank_target_cells": 2,
    }
    assert payload["untouched_input_baseline"] | {"by_instruction_level": None} == {
        "by_instruction_level": None,
        "cell_accuracy": 0.875,
        "cells": 8,
        "correct_cells": 7,
        "pass_rate": 0.5,
        "passed_tasks": 1,
        "tasks": 2,
    }

    dates = payload["date_and_time"]
    assert dates["golden_target_cells"] | {"by_instruction_level": None} == {
        "by_instruction_level": None,
        "date_or_datetime_cells": 1,
        "non_midnight_datetime_cells": 1,
        "tasks_with_date_or_datetime_cells": 1,
        "tasks_with_non_midnight_datetime_cells": 1,
        "tasks_with_time_cells": 1,
        "time_cells": 1,
    }
    assert dates["initial_workbooks"] | {"by_instruction_level": None} == {
        "by_instruction_level": None,
        "date_like_text_cells": 1,
        "tasks_mixing_typed_dates_and_date_like_text": 1,
        "tasks_mixing_typed_dates_and_non_iso_slash_date_text": 1,
        "tasks_with_date_like_text_cells": 1,
        "tasks_with_typed_date_or_time_cells": 2,
        "typed_date_or_time_cells": 3,
    }
    assert dates["workbook_date_systems"] == {
        "golden": {"1900": 1, "1904": 1},
        "initial": {"1900": 1, "1904": 1},
    }

    secrets = (
        "private-alpha",
        "private-beta",
        "private alpha instruction",
        "private beta instruction",
        "Secret ledger",
        "Secret output",
        "UNCHANGED_ANSWER_SECRET",
        "CLEAR_THIS_ANSWER_SECRET",
        "OUTSIDE_TARGET_SECRET",
        "PRIVATE_DYNAMIC_EXTENT",
        "A1",
        "B1",
    )
    assert all(secret not in rendered for secret in secrets)
    assert all(all(isinstance(item, str) for item in items) for items in _lists_in(payload))
    assert json.loads(rendered) == payload


def _public_dataset_dir() -> Path:
    configured = os.environ.get("FORMULABENCH_DATASET_DIR")
    if configured:
        return Path(configured)
    return Path(__file__).resolve().parents[1] / "data" / "spreadsheetbench_verified_400"


@pytest.mark.filterwarnings(
    "ignore:Data Validation extension is not supported and will be removed:UserWarning"
)
def test_public_dataset_reproduces_checked_in_aggregate_audit() -> None:
    dataset_dir = _public_dataset_dir()
    if not (dataset_dir / "dataset.json").exists():
        pytest.skip("SpreadsheetBench Verified 400 is not downloaded")

    actual = analyse(dataset_dir)
    committed_path = (
        Path(__file__).resolve().parents[1] / "experiments" / "public_benchmark_audit.json"
    )
    committed_text = committed_path.read_text(encoding="utf-8")

    assert actual == json.loads(committed_text)
    assert canonical_json(actual) == committed_text
    assert actual["scope"]["instruction_levels"]["cell_level"]["tasks"] == 275
    assert actual["scope"]["instruction_levels"]["sheet_level"]["tasks"] == 125
    assert actual["targets"]["finite_tasks"] == 399
    assert actual["targets"]["dynamic_tasks"] == 1
    assert actual["targets"]["finite_cells"] == 297_854
    assert actual["targets"]["largest_finite_task_cells"] == 104_110

    sheet_golden = actual["golden_targets"]["by_instruction_level"]["sheet_level"]
    assert sheet_golden["blank_cells"] == 134_459
    assert sheet_golden["blank_cell_fraction"] == 0.4754
    assert sheet_golden["tasks_with_blank_cells"] == 88

    sheet_initial = actual["initial_target_state"]["by_instruction_level"]["sheet_level"]
    assert sheet_initial["tasks_with_initial_nonblank_target_cells"] == 98
    assert sheet_initial["tasks_with_correct_nonblank_prefill_and_remaining_changes"] == 86
    assert sheet_initial["tasks_requiring_nonblank_cell_clearing"] == 43
    assert sheet_initial["nonblank_cells_requiring_clearing"] == 100_752

    sheet_baseline = actual["untouched_input_baseline"]["by_instruction_level"]["sheet_level"]
    assert sheet_baseline["cell_accuracy"] == 0.3598
    assert sheet_baseline["passed_tasks"] == 1

    dates = actual["date_and_time"]
    assert dates["golden_target_cells"]["tasks_with_date_or_datetime_cells"] == 37
    assert dates["golden_target_cells"]["date_or_datetime_cells"] == 10_319
    assert dates["golden_target_cells"]["tasks_with_time_cells"] == 4
    assert dates["golden_target_cells"]["tasks_with_non_midnight_datetime_cells"] == 4
    assert dates["initial_workbooks"]["tasks_mixing_typed_dates_and_date_like_text"] == 12
    assert dates["initial_workbooks"]["tasks_mixing_typed_dates_and_non_iso_slash_date_text"] == 7
    assert dates["workbook_date_systems"] == {
        "golden": {"1900": 400},
        "initial": {"1900": 400},
    }
