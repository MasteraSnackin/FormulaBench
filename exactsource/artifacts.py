"""Durable, judge-compatible output artefacts.

The runtime writes only beneath its configured output directory. Existing
unrelated files are left alone; files for the current task set are replaced
atomically so a failed write cannot leave a half-written workbook or JSONL.
"""

from __future__ import annotations

import hashlib
import json
import os
import shutil
import tempfile
from collections.abc import Iterable, Mapping
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import openpyxl

from exactsource.config import TRACE_TEXT_LIMIT

# The organiser explicitly permits truncating workbook serialisation. That data is
# embedded in the prompt. Model responses and executed tool inputs remain complete
# so judges can reconstruct exactly which plan or Python transform was applied.
_TRACE_TEXT_FIELDS = ("prompt",)
_PUBLISHED_FILE_MODE = 0o644
_PUBLISHED_DIRECTORY_MODE = 0o755
_MIDDLE_TRUNCATION_MARKER = "\n...[MIDDLE TRUNCATED]...\n"


class ArtifactError(RuntimeError):
    """Raised when an output artefact is unsafe, missing or malformed."""


@dataclass(frozen=True, slots=True)
class Prediction:
    """One evaluator prediction entry."""

    id: str
    output: str
    status: str

    def as_dict(self) -> dict[str, str]:
        return {"id": self.id, "output": self.output, "status": self.status}


@dataclass(frozen=True, slots=True)
class OutputLayout:
    """All paths created by one ExactSource run."""

    root: Path

    @property
    def outputs_dir(self) -> Path:
        return self.root / "outputs"

    @property
    def traces_dir(self) -> Path:
        return self.root / "traces"

    @property
    def predictions_path(self) -> Path:
        return self.root / "predictions.jsonl"

    @property
    def run_metrics_path(self) -> Path:
        return self.root / "run_metrics.json"

    @property
    def log_path(self) -> Path:
        return self.root / "run.log"

    def output_path(self, task_id: str) -> Path:
        return self.outputs_dir / f"{safe_task_id(task_id)}.xlsx"

    def trace_path(self, task_id: str) -> Path:
        return self.traces_dir / f"{safe_task_id(task_id)}.jsonl"

    def relative_output(self, task_id: str) -> str:
        return f"outputs/{safe_task_id(task_id)}.xlsx"


class TraceRecorder:
    """Collect one task's ordered trace records before an atomic write."""

    def __init__(self, task_id: str, *, text_limit: int = TRACE_TEXT_LIMIT) -> None:
        self.task_id = safe_task_id(task_id)
        self.text_limit = text_limit
        self._records: list[dict[str, Any]] = []

    @property
    def records(self) -> tuple[dict[str, Any], ...]:
        return tuple(self._records)

    def record(self, record: Mapping[str, Any] | None = None, /, **fields: Any) -> None:
        item = dict(record or {})
        item.update(fields)
        item.setdefault("schema_version", 1)
        item.setdefault("task_id", self.task_id)
        item.setdefault("step", len(self._records) + 1)
        self._records.append(truncate_trace_record(item, self.text_limit))

    add = record

    def update_last(self, **fields: Any) -> None:
        """Attach post-call tool evidence to the most recent model-call record."""

        if not self._records:
            raise ArtifactError("cannot update an empty trace")
        item = dict(self._records[-1])
        item.update(fields)
        self._records[-1] = truncate_trace_record(item, self.text_limit)


def safe_task_id(task_id: str) -> str:
    """Return a task id that is safe to use as a single filename component."""

    if not isinstance(task_id, str) or not task_id or task_id != task_id.strip():
        raise ArtifactError(f"unsafe task id: {task_id!r}")
    if (
        task_id in {".", ".."}
        or "/" in task_id
        or "\\" in task_id
        or len(task_id) > 160
        or any(ord(character) < 32 for character in task_id)
    ):
        raise ArtifactError(f"unsafe task id: {task_id!r}")
    return task_id


def prepare_output(out_dir: Path) -> OutputLayout:
    """Create required directories without deleting existing output files."""

    root = Path(out_dir)
    _ensure_directory(root)
    layout = OutputLayout(root=root)
    _ensure_directory(layout.outputs_dir)
    _ensure_directory(layout.traces_dir)
    return layout


def atomic_write_text(path: Path, text: str) -> None:
    """Replace a UTF-8 text file using a temporary sibling and ``os.replace``."""

    path = Path(path)
    _ensure_directory(path.parent)
    temporary = _temporary_sibling(path)
    try:
        with temporary.open("w", encoding="utf-8", newline="\n") as handle:
            handle.write(text)
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(temporary, path)
        os.chmod(path, _PUBLISHED_FILE_MODE)
    finally:
        temporary.unlink(missing_ok=True)


def atomic_write_jsonl(path: Path, records: Iterable[Mapping[str, Any]]) -> None:
    lines = [
        json.dumps(dict(record), ensure_ascii=False, separators=(",", ":")) for record in records
    ]
    atomic_write_text(path, "".join(f"{line}\n" for line in lines))


def atomic_copy_workbook(source: Path, destination: Path) -> None:
    """Copy a workbook only if the completed temporary copy is readable."""

    source = Path(source)
    destination = Path(destination)
    _ensure_directory(destination.parent)
    validate_workbook(source)
    temporary = _temporary_sibling(destination, suffix=".xlsx")
    try:
        shutil.copyfile(source, temporary)
        validate_workbook(temporary)
        os.replace(temporary, destination)
        os.chmod(destination, _PUBLISHED_FILE_MODE)
    finally:
        temporary.unlink(missing_ok=True)


def promote_workbook(temporary: Path, destination: Path) -> None:
    """Validate and atomically promote a solver-produced workbook."""

    temporary = Path(temporary)
    destination = Path(destination)
    validate_workbook(temporary)
    _ensure_directory(destination.parent)
    os.replace(temporary, destination)
    os.chmod(destination, _PUBLISHED_FILE_MODE)


def validate_workbook(path: Path) -> None:
    """Check that ``path`` is a non-empty workbook openpyxl can read."""

    path = Path(path)
    if not path.is_file() or path.is_symlink():
        raise ArtifactError(f"workbook is missing or unsafe: {path}")
    if path.stat().st_size == 0:
        raise ArtifactError(f"workbook is empty: {path}")
    try:
        workbook = openpyxl.load_workbook(
            path,
            read_only=True,
            data_only=False,
            keep_links=False,
        )
        if not workbook.sheetnames:
            raise ArtifactError(f"workbook has no worksheets: {path}")
        workbook.close()
    except ArtifactError:
        raise
    except Exception as error:  # openpyxl exposes several parser exception types
        raise ArtifactError(
            f"unreadable workbook {path}: {type(error).__name__}: {error}"
        ) from error


def write_predictions(layout: OutputLayout, predictions: Iterable[Prediction]) -> None:
    atomic_write_jsonl(
        layout.predictions_path, (prediction.as_dict() for prediction in predictions)
    )


def write_trace(layout: OutputLayout, task_id: str, recorder: TraceRecorder) -> None:
    if recorder.task_id != safe_task_id(task_id):
        raise ArtifactError("trace recorder belongs to a different task")
    atomic_write_jsonl(layout.trace_path(task_id), recorder.records)


def read_jsonl(path: Path) -> list[dict[str, Any]]:
    records: list[dict[str, Any]] = []
    try:
        with Path(path).open(encoding="utf-8") as handle:
            for line_number, line in enumerate(handle, start=1):
                if not line.strip():
                    continue
                value = json.loads(line)
                if not isinstance(value, dict):
                    raise ArtifactError(f"{path}:{line_number} is not a JSON object")
                records.append(value)
    except ArtifactError:
        raise
    except (OSError, json.JSONDecodeError) as error:
        raise ArtifactError(f"cannot read JSONL {path}: {error}") from error
    return records


def validate_run(
    layout: OutputLayout,
    expected_task_ids: Iterable[str],
) -> tuple[Prediction, ...]:
    """Verify the final evaluator contract, including order and readability."""

    expected = tuple(safe_task_id(task_id) for task_id in expected_task_ids)
    if len(set(expected)) != len(expected):
        raise ArtifactError("expected task ids contain duplicates")

    raw_predictions = read_jsonl(layout.predictions_path)
    if len(raw_predictions) != len(expected):
        raise ArtifactError(
            f"predictions count is {len(raw_predictions)}; expected {len(expected)}"
        )

    predictions: list[Prediction] = []
    for index, (raw, task_id) in enumerate(zip(raw_predictions, expected, strict=True), start=1):
        if set(raw) != {"id", "output", "status"}:
            raise ArtifactError(f"prediction {index} has unexpected fields")
        if raw["id"] != task_id:
            raise ArtifactError(f"prediction {index} id is {raw['id']!r}; expected {task_id!r}")
        output = raw["output"]
        status = raw["status"]
        if output != layout.relative_output(task_id):
            raise ArtifactError(f"prediction {task_id!r} has unexpected output path")
        if not isinstance(status, str) or not status:
            raise ArtifactError(f"prediction {task_id!r} has an invalid status")

        validate_workbook(layout.output_path(task_id))
        trace_path = layout.trace_path(task_id)
        if not trace_path.is_file() or trace_path.is_symlink():
            raise ArtifactError(f"trace is missing or unsafe: {trace_path}")
        trace_records = read_jsonl(trace_path)
        if not trace_records:
            if status.startswith("error"):
                predictions.append(Prediction(task_id, output, status))
                continue
            raise ArtifactError(f"successful task has an empty trace: {trace_path}")
        previous_step = 0
        for trace_index, trace in enumerate(trace_records, start=1):
            if trace.get("task_id", task_id) != task_id:
                raise ArtifactError(
                    f"trace {task_id!r} record {trace_index} names a different task"
                )
            step = trace.get("step")
            if not isinstance(step, int) or isinstance(step, bool) or step <= previous_step:
                raise ArtifactError(f"trace {task_id!r} steps are not strictly increasing")
            previous_step = step
            if not isinstance(trace.get("model"), str) or not trace["model"]:
                raise ArtifactError(f"trace {task_id!r} record {trace_index} has no model")
        predictions.append(Prediction(task_id, output, status))
    return tuple(predictions)


def truncate_trace_record(record: Mapping[str, Any], limit: int) -> dict[str, Any]:
    if limit < 1:
        raise ValueError("trace text limit must be positive")
    result = dict(record)
    for field in _TRACE_TEXT_FIELDS:
        value = result.get(field)
        if isinstance(value, str):
            serialised = value
            encoding = "text"
        elif value is not None:
            try:
                serialised = json.dumps(
                    value,
                    ensure_ascii=False,
                    separators=(",", ":"),
                    sort_keys=True,
                    default=str,
                )
            except (TypeError, ValueError):
                serialised = str(value)
            encoding = "json"
        else:
            continue
        if len(serialised) > limit:
            if limit > len(_MIDDLE_TRUNCATION_MARKER):
                retained = limit - len(_MIDDLE_TRUNCATION_MARKER)
                head = (retained + 1) // 2
                tail = retained // 2
                result[field] = (
                    serialised[:head]
                    + _MIDDLE_TRUNCATION_MARKER
                    + (serialised[-tail:] if tail else "")
                )
            else:
                head = (limit + 1) // 2
                tail = limit // 2
                result[field] = serialised[:head] + (serialised[-tail:] if tail else "")
            result[f"{field}_truncated"] = True
            result[f"{field}_truncation"] = "middle"
            result[f"{field}_original_chars"] = len(serialised)
            result[f"{field}_sha256"] = hashlib.sha256(serialised.encode("utf-8")).hexdigest()
            result[f"{field}_encoding"] = encoding
    return result


def temporary_workbook_path(layout: OutputLayout, task_id: str) -> Path:
    """Reserve a unique same-filesystem path for one task's working copy."""

    destination = layout.output_path(task_id)
    return _temporary_sibling(destination, suffix=".xlsx")


def _temporary_sibling(path: Path, *, suffix: str | None = None) -> Path:
    descriptor, name = tempfile.mkstemp(
        prefix=f".{path.stem}.",
        suffix=suffix if suffix is not None else f"{path.suffix}.tmp",
        dir=path.parent,
    )
    os.close(descriptor)
    return Path(name)


def _ensure_directory(path: Path) -> None:
    if path.is_symlink():
        raise ArtifactError(f"refusing symlinked output directory: {path}")
    if path.exists() and not path.is_dir():
        raise ArtifactError(f"output path is not a directory: {path}")
    existed = path.exists()
    path.mkdir(parents=True, exist_ok=True)
    if not existed:
        os.chmod(path, _PUBLISHED_DIRECTORY_MODE)
