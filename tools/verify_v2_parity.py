#!/usr/bin/env python3
"""Verify an ExactSource v2 run against a retained, credential-free reference.

This tool rebuilds each task's initial prompt and replays already-recorded plans. It never reads a
golden workbook, calls a model provider, or writes beneath either input root.
"""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import os
import re
import sys
import tempfile
import zipfile
from collections.abc import Mapping, Sequence
from pathlib import Path
from typing import Any

# Reference v2 workbooks were serialised through openpyxl's stdlib XML path.
# openpyxl reads this switch at import time, so set it before importing any
# ExactSource module (several of which import openpyxl transitively).
os.environ["OPENPYXL_LXML"] = "False"

# ``python tools/verify_v2_parity.py`` puts tools/, rather than the repository
# root, on sys.path.  Add the vendored package root explicitly.
PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from exactsource.context import build_context  # noqa: E402
from exactsource.contracts import SolvePlan, TaskSpec  # noqa: E402
from exactsource.dataset import load_tasks  # noqa: E402
from exactsource.plans import apply_operations  # noqa: E402
from exactsource.prompts import build_messages  # noqa: E402
from exactsource.sandbox import run_transform  # noqa: E402
from formulabench.v2 import _install_exactsource_child_serializer  # noqa: E402

_install_exactsource_child_serializer()

_MIDDLE_TRUNCATION_MARKER = "\n...[MIDDLE TRUNCATED]...\n"
_CORE_PROPERTIES_MEMBER = "docProps/core.xml"
_CORE_TIMESTAMP_RES = (
    (
        "dcterms:created",
        re.compile(rb"(?P<open><dcterms:created(?:\s[^>]*)?>)[^<]*(?P<close></dcterms:created>)"),
    ),
    (
        "dcterms:modified",
        re.compile(rb"(?P<open><dcterms:modified(?:\s[^>]*)?>)[^<]*(?P<close></dcterms:modified>)"),
    ),
)
_NORMALISED_TIMESTAMP_VALUE = b"__FORMULABENCH_CORE_TIMESTAMP__"


class ParityError(RuntimeError):
    """A reference artefact does not match a fresh deterministic reconstruction."""


def _pillow_is_available() -> bool:
    """Return whether Pillow can affect openpyxl's workbook serialisation path."""

    try:
        return importlib.util.find_spec("PIL") is not None
    except (AttributeError, ImportError, ValueError) as exc:
        raise ParityError(f"cannot determine whether Pillow is installed: {exc}") from exc


def _require_canonical_optional_dependencies() -> None:
    """Fail closed when optional packages differ from the canonical verifier image."""

    if _pillow_is_available():
        raise ParityError(
            "Pillow is installed, but the canonical v2 parity environment excludes it; "
            "run scripts/verify_v2_migration.sh in Docker"
        )


def _sha256_text(value: str) -> str:
    return hashlib.sha256(value.encode("utf-8")).hexdigest()


def _read_jsonl(path: Path) -> list[dict[str, Any]]:
    records: list[dict[str, Any]] = []
    try:
        with path.open("r", encoding="utf-8") as handle:
            for line_number, line in enumerate(handle, start=1):
                if not line.strip():
                    continue
                value = json.loads(line)
                if not isinstance(value, dict):
                    raise ParityError(f"{path}:{line_number} is not a JSON object")
                records.append(value)
    except ParityError:
        raise
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise ParityError(f"cannot read JSONL {path}: {exc}") from exc
    return records


def _is_relative_to(path: Path, root: Path) -> bool:
    try:
        path.relative_to(root)
    except ValueError:
        return False
    return True


def _reference_file(root: Path, relative: Path, *, label: str) -> Path:
    if relative.is_absolute() or any(part in {"", ".", ".."} for part in relative.parts):
        raise ParityError(f"unsafe {label} path: {relative}")
    candidate = root / relative
    try:
        resolved = candidate.resolve(strict=True)
    except (FileNotFoundError, OSError) as exc:
        raise ParityError(f"missing {label}: {candidate}") from exc
    if not _is_relative_to(resolved, root):
        raise ParityError(f"{label} escapes the reference root: {candidate}")
    if not resolved.is_file():
        raise ParityError(f"{label} is not a regular file: {candidate}")
    return resolved


def _prediction_map(reference_root: Path) -> dict[str, dict[str, Any]]:
    predictions_path = _reference_file(
        reference_root,
        Path("predictions.jsonl"),
        label="predictions file",
    )
    predictions: dict[str, dict[str, Any]] = {}
    for position, record in enumerate(_read_jsonl(predictions_path), start=1):
        task_id = record.get("id")
        if not isinstance(task_id, str) or not task_id:
            raise ParityError(f"prediction {position} has no valid id")
        if task_id in predictions:
            raise ParityError(f"duplicate reference prediction id: {task_id}")
        status = record.get("status")
        if not isinstance(status, str) or not status:
            raise ParityError(f"prediction {task_id!r} has no valid status")
        expected_output = f"outputs/{task_id}.xlsx"
        if record.get("output") != expected_output:
            raise ParityError(
                f"prediction {task_id!r} output is not the expected {expected_output!r}"
            )
        predictions[task_id] = record
    if not predictions:
        raise ParityError("reference predictions are empty")
    return predictions


def _middle_truncate(value: str, limit: int) -> str:
    if limit > len(_MIDDLE_TRUNCATION_MARKER):
        retained = limit - len(_MIDDLE_TRUNCATION_MARKER)
        head = (retained + 1) // 2
        tail = retained // 2
        return value[:head] + _MIDDLE_TRUNCATION_MARKER + (value[-tail:] if tail else "")
    head = (limit + 1) // 2
    tail = limit // 2
    return value[:head] + (value[-tail:] if tail else "")


def _verify_initial_evidence(task: TaskSpec, first_trace: Mapping[str, Any]) -> None:
    context = build_context(task)
    expected_context = {
        "original_chars": context.original_chars,
        "emitted_chars": len(context.text),
        "truncated": context.truncated,
        "sha256": context.sha256,
    }
    if first_trace.get("context") != expected_context:
        raise ParityError(
            "context metadata mismatch: "
            f"expected {expected_context!r}, found {first_trace.get('context')!r}"
        )

    messages = build_messages(task, context)
    full_prompt = json.dumps(messages, ensure_ascii=False, separators=(",", ":"))
    stored_prompt = first_trace.get("prompt")
    if not isinstance(stored_prompt, str):
        raise ParityError("first trace record has no string prompt")

    if first_trace.get("prompt_truncated") is True:
        expected_digest = _sha256_text(full_prompt)
        metadata = {
            "prompt_truncation": "middle",
            "prompt_original_chars": len(full_prompt),
            "prompt_sha256": expected_digest,
            "prompt_encoding": "text",
        }
        differences = {
            key: (expected, first_trace.get(key))
            for key, expected in metadata.items()
            if first_trace.get(key) != expected
        }
        if differences:
            raise ParityError(f"full prompt truncation metadata mismatch: {differences!r}")
        expected_retained = _middle_truncate(full_prompt, len(stored_prompt))
        if stored_prompt != expected_retained:
            raise ParityError("retained middle-truncated prompt content mismatch")
        return

    if stored_prompt != full_prompt:
        raise ParityError(
            "initial full prompt mismatch: "
            f"expected sha256={_sha256_text(full_prompt)}, "
            f"found sha256={_sha256_text(stored_prompt)}"
        )


def _normalise_zip_member(name: str, payload: bytes) -> bytes:
    """Normalise only openpyxl's runtime-generated core timestamp values."""

    if name != _CORE_PROPERTIES_MEMBER:
        return payload
    normalised = payload
    for element_name, expression in _CORE_TIMESTAMP_RES:
        matches = tuple(expression.finditer(normalised))
        if len(matches) > 1:
            raise ParityError(f"docProps/core.xml has multiple {element_name} elements")
        if not matches:
            continue
        match = matches[0]
        normalised = (
            normalised[: match.start()]
            + match.group("open")
            + _NORMALISED_TIMESTAMP_VALUE
            + match.group("close")
            + normalised[match.end() :]
        )
    return normalised


def _zip_inventory(path: Path) -> tuple[tuple[str, bytes], ...]:
    try:
        with zipfile.ZipFile(path, "r") as archive:
            infos = archive.infolist()
            names = [info.filename for info in infos]
            if len(names) != len(set(names)):
                raise ParityError(f"workbook ZIP has duplicate member names: {path}")
            return tuple(
                (info.filename, _normalise_zip_member(info.filename, archive.read(info)))
                for info in sorted(infos, key=lambda item: item.filename)
            )
    except ParityError:
        raise
    except (OSError, zipfile.BadZipFile, RuntimeError) as exc:
        raise ParityError(f"cannot read workbook ZIP {path}: {exc}") from exc


def zip_content_digest(path: Path) -> str:
    """Hash exact ZIP contents except headers and the two core timestamps."""

    digest = hashlib.sha256()
    for name, payload in _zip_inventory(path):
        encoded_name = name.encode("utf-8")
        digest.update(len(encoded_name).to_bytes(8, "big"))
        digest.update(encoded_name)
        digest.update(len(payload).to_bytes(8, "big"))
        digest.update(payload)
    return digest.hexdigest()


def _workbook_difference(expected: Path, actual: Path) -> str | None:
    expected_inventory = dict(_zip_inventory(expected))
    actual_inventory = dict(_zip_inventory(actual))
    if expected_inventory == actual_inventory:
        return None

    expected_names = set(expected_inventory)
    actual_names = set(actual_inventory)
    missing = sorted(expected_names - actual_names)
    unexpected = sorted(actual_names - expected_names)
    if missing or unexpected:
        return f"ZIP members differ: missing={missing!r}, unexpected={unexpected!r}"
    for name in sorted(expected_names):
        expected_payload = expected_inventory[name]
        actual_payload = actual_inventory[name]
        if expected_payload != actual_payload:
            return (
                f"ZIP member content differs at {name!r}: "
                f"reference_sha256={hashlib.sha256(expected_payload).hexdigest()}, "
                f"replay_sha256={hashlib.sha256(actual_payload).hexdigest()}"
            )
    return "workbook ZIP contents differ"


def _validate_plan(raw: object) -> SolvePlan:
    if not isinstance(raw, Mapping):
        raise ParityError("accepted terminal trace has no SolvePlan object in tool_input")
    try:
        return SolvePlan.model_validate(dict(raw))
    except Exception as exc:
        raise ParityError(f"accepted terminal tool_input is not a valid SolvePlan: {exc}") from exc


def _replay_accepted(
    task: TaskSpec,
    terminal_trace: Mapping[str, Any],
    reference_output: Path,
) -> None:
    if terminal_trace.get("plan_status") != "accepted":
        raise ParityError("successful prediction has no accepted terminal plan")
    plan = _validate_plan(terminal_trace.get("tool_input"))
    with tempfile.TemporaryDirectory(prefix=f"formulabench-parity-{task.id}-") as directory:
        replay_output = Path(directory) / "replayed.xlsx"
        if plan.route == "operations":
            apply_operations(plan, task, task.init_xlsx, replay_output)
        else:
            if task.is_cell_level:
                raise ParityError("cell-level accepted plan unexpectedly uses the Python route")
            assert plan.python_code is not None
            run_transform(plan.python_code, task.init_xlsx, replay_output)

        difference = _workbook_difference(reference_output, replay_output)
        if difference is not None:
            raise ParityError(
                f"accepted-plan replay mismatch ({difference}); "
                f"reference_zip_digest={zip_content_digest(reference_output)}, "
                f"replay_zip_digest={zip_content_digest(replay_output)}"
            )


def _verify_fallback(task: TaskSpec, reference_output: Path) -> None:
    try:
        init_bytes = task.init_xlsx.read_bytes()
        output_bytes = reference_output.read_bytes()
    except OSError as exc:
        raise ParityError(f"cannot read fallback workbook bytes: {exc}") from exc
    if output_bytes != init_bytes:
        raise ParityError(
            "failure output is not byte-identical to init: "
            f"init_sha256={hashlib.sha256(init_bytes).hexdigest()}, "
            f"output_sha256={hashlib.sha256(output_bytes).hexdigest()}"
        )


def _selected_tasks(tasks: Sequence[TaskSpec], ids_argument: str | None) -> list[TaskSpec]:
    if ids_argument is None:
        return list(tasks)
    requested = {item.strip() for item in ids_argument.split(",") if item.strip()}
    if not requested:
        raise ParityError("--ids must contain at least one task id")
    available = {task.id for task in tasks}
    unknown = sorted(requested - available)
    if unknown:
        raise ParityError(f"--ids names unknown dataset tasks: {unknown!r}")
    return [task for task in tasks if task.id in requested]


def _validate_prediction_coverage(
    tasks: Sequence[TaskSpec],
    selected: Sequence[TaskSpec],
    predictions: Mapping[str, Mapping[str, Any]],
    *,
    subset: bool,
) -> None:
    """Require a manifest-ordered full reference or an exact selected subset."""

    manifest_ids = [task.id for task in tasks]
    selected_ids = [task.id for task in selected]
    prediction_ids = list(predictions)
    allowed_orders = [manifest_ids]
    if subset and selected_ids != manifest_ids:
        allowed_orders.append(selected_ids)
    if prediction_ids in allowed_orders:
        return

    if subset:
        raise ParityError(
            "reference prediction IDs must match either the full manifest or the selected "
            f"subset in manifest order: found={prediction_ids!r}, selected={selected_ids!r}, "
            f"manifest={manifest_ids!r}"
        )
    raise ParityError(
        "full-run reference prediction IDs must exactly match the dataset manifest in order: "
        f"found={prediction_ids!r}, expected={manifest_ids!r}"
    )


def verify(dataset_dir: Path, reference_run: Path, ids_argument: str | None = None) -> int:
    """Run all parity checks and return a process-style status code."""

    try:
        _require_canonical_optional_dependencies()
        reference_root = Path(reference_run).resolve(strict=True)
        if not reference_root.is_dir():
            raise ParityError(f"reference run is not a directory: {reference_run}")
        all_tasks = load_tasks(Path(dataset_dir))
        if not all_tasks:
            raise ParityError("dataset contains no tasks")
        tasks = _selected_tasks(all_tasks, ids_argument)
        predictions = _prediction_map(reference_root)
        _validate_prediction_coverage(
            all_tasks,
            tasks,
            predictions,
            subset=ids_argument is not None,
        )
    except Exception as exc:
        print(f"parity setup failed: {exc}", file=sys.stderr)
        print("checked=0 accepted=0 fallback=0 mismatches=1")
        return 1

    accepted = 0
    fallback = 0
    mismatches = 0
    for task in tasks:
        prediction = predictions.get(task.id)
        if prediction is None:
            mismatches += 1
            print(f"[{task.id}] mismatch: reference prediction is missing", file=sys.stderr)
            continue

        is_accepted = prediction["status"] == "ok"
        if is_accepted:
            accepted += 1
        else:
            fallback += 1

        try:
            trace_path = _reference_file(
                reference_root,
                Path("traces") / f"{task.id}.jsonl",
                label=f"trace for task {task.id}",
            )
            trace_records = _read_jsonl(trace_path)
            if not trace_records:
                raise ParityError("reference trace is empty")
            _verify_initial_evidence(task, trace_records[0])

            output_path = _reference_file(
                reference_root,
                Path("outputs") / f"{task.id}.xlsx",
                label=f"output for task {task.id}",
            )
            if is_accepted:
                _replay_accepted(task, trace_records[-1], output_path)
            else:
                _verify_fallback(task, output_path)
        except Exception as exc:
            mismatches += 1
            print(f"[{task.id}] mismatch: {exc}", file=sys.stderr)

    print(f"checked={len(tasks)} accepted={accepted} fallback={fallback} mismatches={mismatches}")
    return 1 if mismatches else 0


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Verify ExactSource v2 prompt, fallback, and replay parity."
    )
    parser.add_argument("--dataset-dir", required=True, type=Path)
    parser.add_argument("--reference-run", required=True, type=Path)
    parser.add_argument("--ids", help="optional comma-separated task ids")
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    arguments = _parser().parse_args(argv)
    return verify(arguments.dataset_dir, arguments.reference_run, arguments.ids)


if __name__ == "__main__":
    raise SystemExit(main())
