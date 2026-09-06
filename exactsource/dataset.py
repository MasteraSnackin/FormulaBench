"""Safe loading of the SpreadsheetBench judge input layout."""

from __future__ import annotations

import json
import os
from collections.abc import Mapping
from pathlib import Path
from typing import Any

import openpyxl

from exactsource.contracts import QualifiedRange, TaskSpec
from exactsource.ranges import RangeSyntaxError, parse_answer_ranges


class DatasetError(ValueError):
    """Raised when judge input is missing, ambiguous or unsafe."""


def _is_relative_to(path: Path, parent: Path) -> bool:
    try:
        path.relative_to(parent)
    except ValueError:
        return False
    return True


def _resolve_within(path: Path, parent: Path, *, kind: str) -> Path:
    """Resolve symlinks and reject any escape from ``parent``."""

    try:
        resolved = path.resolve(strict=True)
    except (FileNotFoundError, OSError) as exc:
        raise DatasetError(f"{kind} does not exist: {path}") from exc
    if not _is_relative_to(resolved, parent):
        raise DatasetError(f"{kind} escapes the dataset root: {path}")
    return resolved


def _safe_task_id(raw: Any) -> str:
    if isinstance(raw, bool) or not isinstance(raw, (str, int)):
        raise DatasetError(f"task id must be a string or integer, got {type(raw).__name__}")
    task_id = str(raw).strip()
    if not task_id:
        raise DatasetError("task id must not be empty")
    if task_id in {".", ".."} or "/" in task_id or "\\" in task_id:
        raise DatasetError(f"unsafe task id: {task_id!r}")
    if any(ord(character) < 32 for character in task_id):
        raise DatasetError(f"task id contains a control character: {task_id!r}")
    if len(task_id) > 160:
        raise DatasetError(f"task id is too long: {task_id!r}")
    return task_id


def _required_text(record: Mapping[str, Any], field: str, *, task_id: str) -> str:
    value = record.get(field)
    if not isinstance(value, str) or not value.strip():
        raise DatasetError(f"task {task_id}: {field} must be a non-empty string")
    return value.strip()


def _task_directory(root: Path, raw: Any, *, task_id: str) -> tuple[str, Path]:
    if not isinstance(raw, str) or not raw.strip():
        raise DatasetError(f"task {task_id}: spreadsheet_path must be a non-empty string")
    spreadsheet_path = raw.strip()
    if "\\" in spreadsheet_path:
        raise DatasetError(f"task {task_id}: spreadsheet_path must use '/' separators")
    relative = Path(spreadsheet_path)
    if relative.is_absolute() or any(part in {"", ".", ".."} for part in relative.parts):
        raise DatasetError(f"task {task_id}: unsafe spreadsheet_path {spreadsheet_path!r}")
    folder = _resolve_within(root / relative, root, kind=f"task {task_id} directory")
    if not folder.is_dir():
        raise DatasetError(f"task {task_id}: spreadsheet_path is not a directory")
    return spreadsheet_path, folder


def _find_init_workbook(folder: Path, root: Path, *, task_id: str) -> Path:
    # A malformed or adversarial dataset can name a golden workbook so that it
    # also matches ``*init*.xlsx``. Exclude such names explicitly before any
    # workbook is resolved or opened.
    candidates = sorted(
        (item for item in folder.glob("*init*.xlsx") if "golden" not in item.name.casefold()),
        key=lambda item: item.name.casefold(),
    )
    if len(candidates) != 1:
        raise DatasetError(
            f"task {task_id}: expected exactly one *init*.xlsx workbook, found {len(candidates)}"
        )
    candidate = candidates[0]
    if candidate.is_symlink():
        raise DatasetError(f"task {task_id}: init workbook must not be a symbolic link")
    workbook = _resolve_within(candidate, root, kind=f"task {task_id} init workbook")
    if "golden" in workbook.name.casefold():
        raise DatasetError(f"task {task_id}: init workbook resolves to a forbidden filename")
    if not workbook.is_file():
        raise DatasetError(f"task {task_id}: init workbook is not a regular file")
    return workbook


def _read_instruction(record: Mapping[str, Any], folder: Path, root: Path, *, task_id: str) -> str:
    prompt_path = folder / "prompt.txt"
    if os.path.lexists(prompt_path):
        resolved = _resolve_within(prompt_path, root, kind=f"task {task_id} prompt")
        if not resolved.is_file():
            raise DatasetError(f"task {task_id}: prompt.txt is not a regular file")
        try:
            prompt = resolved.read_text(encoding="utf-8").strip()
        except UnicodeDecodeError as exc:
            raise DatasetError(f"task {task_id}: prompt.txt is not valid UTF-8") from exc
        if prompt:
            return prompt

    instruction = record.get("instruction")
    if not isinstance(instruction, str) or not instruction.strip():
        raise DatasetError(f"task {task_id}: no non-empty prompt.txt or instruction field")
    return instruction.strip()


def _workbook_sheet_info(workbook: Path, *, task_id: str) -> tuple[tuple[str, ...], str]:
    """Read sheet identity for active fallback and exact operation targeting."""

    try:
        loaded = openpyxl.load_workbook(workbook, read_only=True, data_only=False)
    except Exception as exc:  # openpyxl exposes several format-specific errors
        raise DatasetError(f"task {task_id}: cannot open init workbook: {exc}") from exc
    try:
        active = loaded.active
        if active is None:
            raise DatasetError(f"task {task_id}: init workbook has no active worksheet")
        return tuple(loaded.sheetnames), active.title
    finally:
        loaded.close()


def _canonicalise_answer_sheets(
    ranges: tuple[QualifiedRange, ...],
    workbook_sheets: tuple[str, ...],
) -> tuple[QualifiedRange, ...]:
    """Use an existing tab's exact spelling while preserving future tabs."""

    canonical: list[QualifiedRange] = []
    for target in ranges:
        resolved = target.sheet
        if resolved not in workbook_sheets:
            case_matches = [
                name for name in workbook_sheets if name.casefold() == resolved.casefold()
            ]
            if len(case_matches) == 1:
                resolved = case_matches[0]
            else:
                stripped = resolved.strip().casefold()
                edge_space_matches = [
                    name for name in workbook_sheets if name.strip().casefold() == stripped
                ]
                if len(edge_space_matches) == 1:
                    resolved = edge_space_matches[0]
        canonical.append(QualifiedRange(sheet=resolved, cells=target.cells))
    return tuple(canonical)


def _load_record(record: Mapping[str, Any], root: Path, *, position: int) -> TaskSpec:
    if "id" not in record:
        raise DatasetError(f"task at index {position} has no id")
    task_id = _safe_task_id(record["id"])
    spreadsheet_path, folder = _task_directory(
        root,
        record.get("spreadsheet_path"),
        task_id=task_id,
    )
    init_xlsx = _find_init_workbook(folder, root, task_id=task_id)
    workbook_sheets, active_sheet = _workbook_sheet_info(init_xlsx, task_id=task_id)
    instruction = _read_instruction(record, folder, root, task_id=task_id)
    instruction_type = _required_text(record, "instruction_type", task_id=task_id)
    answer_position = _required_text(record, "answer_position", task_id=task_id)
    answer_sheet_raw = record.get("answer_sheet")
    if answer_sheet_raw is not None and not isinstance(answer_sheet_raw, str):
        raise DatasetError(f"task {task_id}: answer_sheet must be a string or null")
    if answer_sheet_raw is None or not answer_sheet_raw.strip():
        answer_sheet_raw = active_sheet
    try:
        answer_ranges = parse_answer_ranges(answer_position, answer_sheet_raw)
    except RangeSyntaxError as exc:
        raise DatasetError(f"task {task_id}: invalid answer range: {exc}") from exc

    answer_ranges = _canonicalise_answer_sheets(answer_ranges, workbook_sheets)

    data_position = record.get("data_position")
    if data_position is not None:
        if not isinstance(data_position, str):
            raise DatasetError(f"task {task_id}: data_position must be a string or null")
        data_position = data_position.strip() or None

    return TaskSpec(
        id=task_id,
        instruction_type=instruction_type,
        instruction=instruction,
        spreadsheet_path=spreadsheet_path,
        init_xlsx=init_xlsx,
        answer_ranges=answer_ranges,
        data_position=data_position,
    )


def load_tasks(data_dir: Path) -> list[TaskSpec]:
    """Load every task from a judge dataset without touching golden workbooks.

    The only directory search performed is the narrow ``*init*.xlsx`` lookup
    inside each declared task folder.  Every resolved path must remain beneath
    the real dataset root, including through symlinks.
    """

    try:
        root = Path(data_dir).resolve(strict=True)
    except (FileNotFoundError, OSError) as exc:
        raise DatasetError(f"dataset directory does not exist: {data_dir}") from exc
    if not root.is_dir():
        raise DatasetError(f"dataset path is not a directory: {data_dir}")

    dataset_path = _resolve_within(root / "dataset.json", root, kind="dataset.json")
    if not dataset_path.is_file():
        raise DatasetError("dataset.json is not a regular file")
    try:
        payload = json.loads(dataset_path.read_text(encoding="utf-8"))
    except UnicodeDecodeError as exc:
        raise DatasetError("dataset.json is not valid UTF-8") from exc
    except json.JSONDecodeError as exc:
        raise DatasetError(f"dataset.json is invalid JSON: {exc.msg}") from exc
    if not isinstance(payload, list):
        raise DatasetError("dataset.json must contain a JSON array")

    tasks: list[TaskSpec] = []
    seen: set[str] = set()
    for position, raw_record in enumerate(payload):
        if not isinstance(raw_record, Mapping):
            raise DatasetError(f"task at index {position} must be a JSON object")
        task = _load_record(raw_record, root, position=position)
        if task.id in seen:
            raise DatasetError(f"duplicate task id: {task.id}")
        seen.add(task.id)
        tasks.append(task)
    return tasks


# A readable alias for callers that treat the dataset as a single object.
load_dataset = load_tasks
