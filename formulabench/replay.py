"""Credential-free replay of previously stored model responses.

This module repairs the narrow case where a model call completed and its
response was retained in a trace, but workbook materialisation failed.  It
never imports a provider or makes a model call.  Canonical artefacts are
replaced only through the same durable retry journal used by paid retries.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any

from .artifacts import (
    ArtifactLayout,
    DatasetTask,
    ResumeStateError,
    RetryWorkspace,
    _read_resume_jsonl,
    atomic_write_jsonl,
    commit_retry_transaction,
    create_retry_workspace,
    discard_retry_workspace,
    load_resume_state,
    prepare_retry_transaction,
    recover_retry_transactions,
    sha256_file,
)
from .contract import FailureCode, make_prediction_record, write_response_atomic

_REPLAY_FIELD = "replay"
_SOURCE_ERROR = "error:workbook_write_failed"
_SOURCE_FAILURE_CODES = [FailureCode.WORKBOOK_WRITE_FAILED.value]


@dataclass(frozen=True, slots=True)
class ReplayFailure:
    """One stored response that remains invalid under the current writer."""

    task_id: str
    failure_codes: tuple[str, ...]
    error: str | None


@dataclass(frozen=True, slots=True)
class ReplaySummary:
    """Deterministic, manifest-ordered outcome of one local replay pass."""

    total_tasks: int
    eligible_tasks: int
    replayed_tasks: tuple[str, ...]
    still_failed: tuple[ReplayFailure, ...]
    model_calls: int = 0


@dataclass(frozen=True, slots=True)
class _ReplayCandidate:
    task: DatasetTask
    old_prediction: dict[str, Any]
    workspace: RetryWorkspace
    source_trace: dict[str, Any]
    source_trace_sha256: str


def _discard_unprepared(
    layout: ArtifactLayout,
    workspaces: list[RetryWorkspace],
) -> None:
    """Best-effort cleanup that never removes a durable transaction marker."""

    cleanup_error: BaseException | None = None
    for workspace in workspaces:
        if not workspace.directory.exists() or (workspace.directory / "meta.json").exists():
            continue
        try:
            discard_retry_workspace(layout, workspace)
        except BaseException as exc:  # Preserve the first cleanup failure deterministically.
            if cleanup_error is None:
                cleanup_error = exc
    if cleanup_error is not None:
        raise cleanup_error


def _load_source_trace(
    candidate_task: DatasetTask,
    workspace: RetryWorkspace,
    *,
    expected_model: str,
) -> tuple[dict[str, Any], str]:
    """Validate the copied trace as an exact, never-before-replayed source."""

    path = workspace.directory / "old-trace.jsonl"
    records = _read_resume_jsonl(path, code="replay_trace", task_id=candidate_task.id)
    if len(records) != 1:
        raise ResumeStateError("replay_trace_record_count", task_id=candidate_task.id)
    source = records[0]
    if not isinstance(source, dict):
        raise ResumeStateError("replay_trace_record_type", task_id=candidate_task.id)
    step = source.get("step")
    if isinstance(step, bool) or step != 1:
        raise ResumeStateError("replay_trace_step", task_id=candidate_task.id)
    if source.get("model") != expected_model:
        raise ResumeStateError("replay_trace_model", task_id=candidate_task.id)
    if not isinstance(source.get("response"), str):
        raise ResumeStateError("replay_trace_response", task_id=candidate_task.id)
    if source.get("failure_codes") != _SOURCE_FAILURE_CODES:
        raise ResumeStateError("replay_trace_failure_codes", task_id=candidate_task.id)
    if source.get("error") != _SOURCE_ERROR:
        raise ResumeStateError("replay_trace_error", task_id=candidate_task.id)
    if _REPLAY_FIELD in source:
        raise ResumeStateError("replay_already_applied", task_id=candidate_task.id)

    context = source.get("context")
    if not isinstance(context, dict) or not isinstance(context.get("input_sha256"), str):
        raise ResumeStateError("replay_trace_context", task_id=candidate_task.id)
    try:
        input_sha256 = sha256_file(candidate_task.init_xlsx)
        source_trace_sha256 = sha256_file(path)
    except OSError as exc:
        raise ResumeStateError("replay_source_unreadable", task_id=candidate_task.id) from exc
    if context["input_sha256"] != input_sha256:
        raise ResumeStateError("replay_input_hash", task_id=candidate_task.id)
    return dict(source), source_trace_sha256


def _successful_trace(candidate: _ReplayCandidate) -> dict[str, Any]:
    trace = dict(candidate.source_trace)
    trace["error"] = None
    trace["failure_codes"] = []
    trace[_REPLAY_FIELD] = {
        "mode": "stored_response",
        "additional_model_calls": 0,
        "reason": "workbook_write_contract_fix",
        "source_step": 1,
        "source_trace_sha256": candidate.source_trace_sha256,
        "source_error": _SOURCE_ERROR,
        "source_failure_codes": list(_SOURCE_FAILURE_CODES),
    }
    return trace


def replay_stored_responses(
    *,
    layout: ArtifactLayout,
    tasks: list[DatasetTask],
    expected_model: str,
) -> ReplaySummary:
    """Replay exact ``workbook_write_failed`` traces without provider access.

    The selected run must already be complete and valid.  Every eligible trace
    is copied into an isolated retry workspace and preflighted before any
    canonical file is changed.  A response that still fails is reported and
    its scratch workspace is removed; its canonical artefacts remain exact.
    """

    task_ids = [task.id for task in tasks]
    if len(task_ids) != len(set(task_ids)):
        raise ResumeStateError("replay_task_duplicate")

    recover_retry_transactions(layout, tasks, expected_model=expected_model)
    state = load_resume_state(layout, tasks, expected_model=expected_model)
    if state.completed_ids != frozenset(task_ids):
        raise ResumeStateError("replay_incomplete")

    eligible = [
        task
        for task in tasks
        if list(state.predictions[task.id]["failure_codes"]) == _SOURCE_FAILURE_CODES
    ]
    if not eligible:
        return ReplaySummary(
            total_tasks=len(tasks),
            eligible_tasks=0,
            replayed_tasks=(),
            still_failed=(),
        )

    workspaces: list[RetryWorkspace] = []
    candidates: list[_ReplayCandidate] = []
    try:
        for task in eligible:
            workspaces.append(create_retry_workspace(layout, task.id))
        for task, workspace in zip(eligible, workspaces, strict=True):
            source_trace, source_trace_sha256 = _load_source_trace(
                task,
                workspace,
                expected_model=expected_model,
            )
            candidates.append(
                _ReplayCandidate(
                    task=task,
                    old_prediction=dict(state.predictions[task.id]),
                    workspace=workspace,
                    source_trace=source_trace,
                    source_trace_sha256=source_trace_sha256,
                )
            )
    except BaseException:
        _discard_unprepared(layout, workspaces)
        raise

    replayed: list[str] = []
    still_failed: list[ReplayFailure] = []
    try:
        for candidate in candidates:
            write_result = write_response_atomic(
                candidate.task.as_dict(),
                candidate.source_trace["response"],
                candidate.task.init_xlsx,
                candidate.workspace.new_output,
            )
            if not write_result.success:
                failure_codes = tuple(code.value for code in write_result.failure_codes)
                still_failed.append(
                    ReplayFailure(
                        task_id=candidate.task.id,
                        failure_codes=failure_codes,
                        error=write_result.status,
                    )
                )
                discard_retry_workspace(layout, candidate.workspace)
                continue

            atomic_write_jsonl(
                candidate.workspace.new_trace,
                [_successful_trace(candidate)],
            )
            new_prediction = make_prediction_record(
                candidate.task.id,
                Path("outputs") / f"{candidate.task.id}.xlsx",
                write_result,
            )
            prepare_retry_transaction(
                candidate.workspace,
                task=candidate.task,
                old_prediction=candidate.old_prediction,
                new_prediction=new_prediction,
                expected_model=expected_model,
            )
            committed = commit_retry_transaction(
                layout,
                tasks,
                task_id=candidate.task.id,
                expected_model=expected_model,
            )
            if committed != new_prediction:
                raise RuntimeError("replay commit returned an unexpected prediction")
            replayed.append(candidate.task.id)
    except BaseException:
        _discard_unprepared(layout, workspaces)
        raise

    return ReplaySummary(
        total_tasks=len(tasks),
        eligible_tasks=len(candidates),
        replayed_tasks=tuple(replayed),
        still_failed=tuple(still_failed),
    )


__all__ = ["ReplayFailure", "ReplaySummary", "replay_stored_responses"]
