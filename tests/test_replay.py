from __future__ import annotations

import builtins
import copy
import json
from pathlib import Path
from typing import Any

import openpyxl
import pytest

import formulabench.replay as replay_module
from formulabench.artifacts import (
    ArtifactLayout,
    DatasetTask,
    ResumeStateError,
    atomic_copy,
    atomic_write_jsonl,
    load_dataset_manifest,
    load_resume_state,
    recover_retry_transactions,
    sha256_file,
)
from formulabench.constants import MODEL_ID
from formulabench.replay import ReplayFailure, ReplaySummary, replay_stored_responses

_VALID_RESPONSE = '{"cells":[{"sheet":"Inputs","cell":"B2","value":"=SUM(A1:A2)"}]}'


def _dataset(tmp_path: Path, task_ids: tuple[str, ...] = ("task-1",)) -> list[DatasetTask]:
    dataset = tmp_path / "dataset"
    manifest: list[dict[str, object]] = []
    for index, task_id in enumerate(task_ids, start=1):
        task_dir = dataset / "spreadsheet" / task_id
        task_dir.mkdir(parents=True)
        workbook_path = task_dir / f"{index}_{task_id}_init.xlsx"
        workbook = openpyxl.Workbook()
        worksheet = workbook.active
        worksheet.title = "Inputs"
        worksheet["A1"] = index
        worksheet["A2"] = index + 1
        workbook.save(workbook_path)
        workbook.close()
        manifest.append(
            {
                "id": task_id,
                "instruction": "Put the sum of A1:A2 in B2.",
                "spreadsheet_path": f"spreadsheet/{task_id}",
                "answer_position": "B2",
                "answer_sheet": "Inputs",
                "data_position": "Inputs!A1:A2",
            }
        )
    (dataset / "dataset.json").write_text(json.dumps(manifest), encoding="utf-8")
    return load_dataset_manifest(dataset)


def _trace(task: DatasetTask, *, response: str | None = _VALID_RESPONSE) -> dict[str, Any]:
    return {
        "step": 1,
        "model": MODEL_ID,
        "prompt": f"private prompt for {task.id}",
        "response": response,
        "input_tokens": 701,
        "output_tokens": 83,
        "latency_ms": 4567,
        "error": "error:workbook_write_failed",
        "failure_codes": ["workbook_write_failed"],
        "renderer": "pinned-test-renderer",
        "stop_reason": "stop",
        "parse_termination": "stop_sequence",
        "response_format": "strict_json_fallback",
        "response_rejection": None,
        "model_provenance": "configured_base_model",
        "request": {
            "model": MODEL_ID,
            "temperature": 0.0,
            "telemetry": "disabled",
        },
        "context": {
            "characters": 1234,
            "included_cells": 2,
            "omitted_cells": 0,
            "truncated": False,
            "input_sha256": sha256_file(task.init_xlsx),
        },
    }


def _checkpoint(
    layout: ArtifactLayout,
    task: DatasetTask,
    *,
    trace_records: list[dict[str, Any]] | None = None,
    prediction_failure_codes: list[str] | None = None,
) -> dict[str, Any]:
    failure_codes = prediction_failure_codes or ["workbook_write_failed"]
    prediction = {
        "id": task.id,
        "output": f"outputs/{task.id}.xlsx",
        "status": f"error:{','.join(failure_codes)}",
        "failure_codes": failure_codes,
    }
    atomic_copy(task.init_xlsx, layout.outputs / f"{task.id}.xlsx")
    atomic_write_jsonl(
        layout.traces / f"{task.id}.jsonl",
        trace_records if trace_records is not None else [_trace(task)],
    )
    return prediction


def _canonical_bytes(layout: ArtifactLayout, task_ids: tuple[str, ...]) -> dict[str, bytes]:
    paths = [layout.predictions]
    paths.extend(layout.outputs / f"{task_id}.xlsx" for task_id in task_ids)
    paths.extend(layout.traces / f"{task_id}.jsonl" for task_id in task_ids)
    return {path.relative_to(layout.root).as_posix(): path.read_bytes() for path in paths}


def test_replay_success_preserves_model_evidence_and_uses_zero_model_calls(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    tasks = _dataset(tmp_path)
    task = tasks[0]
    layout = ArtifactLayout.initialise(tmp_path / "out")
    original_prediction = _checkpoint(layout, task)
    atomic_write_jsonl(layout.predictions, [original_prediction])
    original_input = task.init_xlsx.read_bytes()
    original_trace_bytes = (layout.traces / f"{task.id}.jsonl").read_bytes()
    original_trace = json.loads(original_trace_bytes)
    original_trace["future_evidence"] = {"nested": [1, {"kept": True}]}
    atomic_write_jsonl(layout.traces / f"{task.id}.jsonl", [original_trace])
    original_trace_sha256 = sha256_file(layout.traces / f"{task.id}.jsonl")

    # Fail immediately if replay ever grows a runtime provider dependency.  A
    # missing key independently proves that this successful path needs no
    # credential lookup.
    monkeypatch.delenv("TINKER_API_KEY", raising=False)
    assert not any(
        name == "provider"
        or getattr(value, "__name__", "") == "formulabench.provider"
        or getattr(value, "__module__", "").startswith("formulabench.provider")
        for name, value in vars(replay_module).items()
    )
    original_import = builtins.__import__

    def reject_provider_import(
        name: str,
        globals: dict[str, Any] | None = None,
        locals: dict[str, Any] | None = None,
        fromlist: tuple[str, ...] = (),
        level: int = 0,
    ) -> Any:
        if name == "formulabench.provider" or name.endswith(".provider"):
            raise AssertionError("stored-response replay must not import a provider")
        return original_import(name, globals, locals, fromlist, level)

    monkeypatch.setattr(builtins, "__import__", reject_provider_import)

    summary = replay_stored_responses(
        layout=layout,
        tasks=tasks,
        expected_model=MODEL_ID,
    )

    assert summary == ReplaySummary(
        total_tasks=1,
        eligible_tasks=1,
        replayed_tasks=("task-1",),
        still_failed=(),
        model_calls=0,
    )
    assert task.init_xlsx.read_bytes() == original_input
    prediction = json.loads(layout.predictions.read_text(encoding="utf-8"))
    assert prediction == {
        "id": "task-1",
        "output": "outputs/task-1.xlsx",
        "status": "ok",
        "failure_codes": [],
    }
    workbook = openpyxl.load_workbook(layout.outputs / "task-1.xlsx", data_only=False)
    assert workbook["Inputs"]["B2"].value == "=SUM(A1:A2)"
    workbook.close()

    replayed_trace = json.loads((layout.traces / "task-1.jsonl").read_text(encoding="utf-8"))
    expected_trace = copy.deepcopy(original_trace)
    expected_trace["error"] = None
    expected_trace["failure_codes"] = []
    expected_trace["replay"] = {
        "mode": "stored_response",
        "additional_model_calls": 0,
        "reason": "workbook_write_contract_fix",
        "source_step": 1,
        "source_trace_sha256": original_trace_sha256,
        "source_error": "error:workbook_write_failed",
        "source_failure_codes": ["workbook_write_failed"],
    }
    assert replayed_trace == expected_trace
    state = load_resume_state(layout, tasks, expected_model=MODEL_ID)
    assert state.successful_ids == frozenset({"task-1"})
    assert not (layout.root / ".retry").exists()

    completed_bytes = _canonical_bytes(layout, ("task-1",))
    second = replay_stored_responses(layout=layout, tasks=tasks, expected_model=MODEL_ID)
    assert second == ReplaySummary(1, 0, (), (), 0)
    assert _canonical_bytes(layout, ("task-1",)) == completed_bytes


def test_replay_that_still_fails_leaves_every_canonical_byte_unchanged(tmp_path: Path) -> None:
    tasks = _dataset(tmp_path)
    task = tasks[0]
    layout = ArtifactLayout.initialise(tmp_path / "out")
    prediction = _checkpoint(layout, task, trace_records=[_trace(task, response='{"cells":[]}')])
    atomic_write_jsonl(layout.predictions, [prediction])
    before = _canonical_bytes(layout, ("task-1",))

    summary = replay_stored_responses(layout=layout, tasks=tasks, expected_model=MODEL_ID)

    assert summary == ReplaySummary(
        total_tasks=1,
        eligible_tasks=1,
        replayed_tasks=(),
        still_failed=(
            ReplayFailure(
                task_id="task-1",
                failure_codes=("missing_cells",),
                error="error:missing_cells",
            ),
        ),
        model_calls=0,
    )
    assert _canonical_bytes(layout, ("task-1",)) == before
    assert not (layout.root / ".retry").exists()


def test_replay_requires_a_complete_validated_resume_before_copying(tmp_path: Path) -> None:
    tasks = _dataset(tmp_path, ("task-1", "task-2"))
    layout = ArtifactLayout.initialise(tmp_path / "out")
    prediction = _checkpoint(layout, tasks[0])
    atomic_write_jsonl(layout.predictions, [prediction])
    before = _canonical_bytes(layout, ("task-1",))

    with pytest.raises(ResumeStateError, match="replay_incomplete") as raised:
        replay_stored_responses(layout=layout, tasks=tasks, expected_model=MODEL_ID)

    assert raised.value.code == "replay_incomplete"
    assert _canonical_bytes(layout, ("task-1",)) == before
    assert not (layout.root / ".retry").exists()


def test_replay_ignores_failures_other_than_the_exact_write_failure(tmp_path: Path) -> None:
    tasks = _dataset(tmp_path)
    task = tasks[0]
    layout = ArtifactLayout.initialise(tmp_path / "out")
    trace = _trace(task, response=None)
    trace["error"] = "model_call_failed:ProviderResponseError"
    trace["failure_codes"] = ["model_call_failed"]
    prediction = _checkpoint(
        layout,
        task,
        trace_records=[trace],
        prediction_failure_codes=["model_call_failed"],
    )
    atomic_write_jsonl(layout.predictions, [prediction])
    before = _canonical_bytes(layout, ("task-1",))

    summary = replay_stored_responses(layout=layout, tasks=tasks, expected_model=MODEL_ID)

    assert summary == ReplaySummary(1, 0, (), (), 0)
    assert _canonical_bytes(layout, ("task-1",)) == before


def test_replay_does_not_select_a_prediction_with_additional_failure_codes(
    tmp_path: Path,
) -> None:
    tasks = _dataset(tmp_path)
    task = tasks[0]
    layout = ArtifactLayout.initialise(tmp_path / "out")
    prediction = _checkpoint(
        layout,
        task,
        prediction_failure_codes=["workbook_write_failed", "missing_cells"],
    )
    atomic_write_jsonl(layout.predictions, [prediction])
    before = _canonical_bytes(layout, ("task-1",))

    summary = replay_stored_responses(layout=layout, tasks=tasks, expected_model=MODEL_ID)

    assert summary == ReplaySummary(1, 0, (), (), 0)
    assert _canonical_bytes(layout, ("task-1",)) == before


@pytest.mark.parametrize(
    ("mutation", "expected_code"),
    [
        ("records", "replay_trace_record_count"),
        ("response", "replay_trace_response"),
        ("failure_codes", "replay_trace_failure_codes"),
        ("error", "replay_trace_error"),
        ("prior_replay", "replay_already_applied"),
        ("context", "replay_trace_context"),
        ("input_hash", "replay_input_hash"),
    ],
)
def test_replay_rejects_inexact_or_tampered_source_trace_before_writing(
    tmp_path: Path,
    mutation: str,
    expected_code: str,
) -> None:
    tasks = _dataset(tmp_path)
    task = tasks[0]
    layout = ArtifactLayout.initialise(tmp_path / "out")
    trace = _trace(task)
    records = [trace]
    if mutation == "records":
        second = copy.deepcopy(trace)
        second["step"] = 2
        records.append(second)
    elif mutation == "response":
        trace["response"] = None
    elif mutation == "failure_codes":
        trace["failure_codes"] = ["missing_cells"]
    elif mutation == "error":
        trace["error"] = "workbook_write_failed:private-detail"
    elif mutation == "prior_replay":
        trace["replay"] = {"mode": "stored_response"}
    elif mutation == "context":
        trace.pop("context")
    elif mutation == "input_hash":
        trace["context"]["input_sha256"] = "0" * 64
    else:  # pragma: no cover - exhaustive parametrisation guard.
        raise AssertionError(mutation)
    prediction = _checkpoint(layout, task, trace_records=records)
    atomic_write_jsonl(layout.predictions, [prediction])
    before = _canonical_bytes(layout, ("task-1",))

    with pytest.raises(ResumeStateError) as raised:
        replay_stored_responses(layout=layout, tasks=tasks, expected_model=MODEL_ID)

    assert raised.value.code == expected_code
    assert _canonical_bytes(layout, ("task-1",)) == before
    assert not (layout.root / ".retry").exists()


def test_all_sources_are_preflighted_before_the_first_canonical_publish(tmp_path: Path) -> None:
    tasks = _dataset(tmp_path, ("task-1", "task-2"))
    layout = ArtifactLayout.initialise(tmp_path / "out")
    predictions = [_checkpoint(layout, tasks[0])]
    bad_trace = _trace(tasks[1])
    bad_trace["context"]["input_sha256"] = "f" * 64
    predictions.append(_checkpoint(layout, tasks[1], trace_records=[bad_trace]))
    atomic_write_jsonl(layout.predictions, predictions)
    before = _canonical_bytes(layout, ("task-1", "task-2"))

    with pytest.raises(ResumeStateError) as raised:
        replay_stored_responses(layout=layout, tasks=tasks, expected_model=MODEL_ID)

    assert raised.value.code == "replay_input_hash"
    assert _canonical_bytes(layout, ("task-1", "task-2")) == before
    assert not (layout.root / ".retry").exists()


def test_prepared_replay_is_recoverable_if_publication_is_interrupted(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    tasks = _dataset(tmp_path)
    task = tasks[0]
    layout = ArtifactLayout.initialise(tmp_path / "out")
    prediction = _checkpoint(layout, task)
    atomic_write_jsonl(layout.predictions, [prediction])
    before = _canonical_bytes(layout, ("task-1",))

    class SimulatedHardStop(BaseException):
        pass

    def stop_before_publication(*args: object, **kwargs: object) -> dict[str, Any]:
        del args, kwargs
        raise SimulatedHardStop

    monkeypatch.setattr(replay_module, "commit_retry_transaction", stop_before_publication)
    with pytest.raises(SimulatedHardStop):
        replay_stored_responses(layout=layout, tasks=tasks, expected_model=MODEL_ID)

    assert _canonical_bytes(layout, ("task-1",)) == before
    assert (layout.root / ".retry" / "task-1" / "meta.json").is_file()

    recover_retry_transactions(layout, tasks, expected_model=MODEL_ID)
    recovered = load_resume_state(layout, tasks, expected_model=MODEL_ID)
    assert recovered.successful_ids == frozenset({"task-1"})
    assert not (layout.root / ".retry").exists()


def test_duplicate_task_selection_is_rejected_before_resume_recovery(tmp_path: Path) -> None:
    tasks = _dataset(tmp_path)
    layout = ArtifactLayout.initialise(tmp_path / "out")
    prediction = _checkpoint(layout, tasks[0])
    atomic_write_jsonl(layout.predictions, [prediction])
    before = _canonical_bytes(layout, ("task-1",))

    with pytest.raises(ResumeStateError) as raised:
        replay_stored_responses(
            layout=layout,
            tasks=[tasks[0], tasks[0]],
            expected_model=MODEL_ID,
        )

    assert raised.value.code == "replay_task_duplicate"
    assert _canonical_bytes(layout, ("task-1",)) == before
