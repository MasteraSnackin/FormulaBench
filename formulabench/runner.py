"""Deterministic inspect, generate, validate and write pipeline."""

from __future__ import annotations

import asyncio
import time
from collections.abc import Mapping
from contextlib import suppress
from dataclasses import dataclass
from pathlib import Path

import openpyxl

from .artifacts import (
    ArtifactLayout,
    DatasetTask,
    RetryWorkspace,
    atomic_copy,
    atomic_write_jsonl,
    commit_retry_transaction,
    create_retry_workspace,
    discard_retry_workspace,
    prepare_retry_transaction,
)
from .constants import (
    DEFAULT_MAX_OUTPUT_TOKENS,
    MODEL_ID,
    MODEL_PROVENANCE,
    NUM_SAMPLES,
    PROVIDER_MAX_RETRIES,
    REASONING_EFFORT,
    RENDERER_ID,
    STOP_TOKEN,
    TEMPERATURE,
    THINKING_MODE,
    TINKER_TELEMETRY_MODE,
    TOKENIZER_REVISION,
    TRANSPORT_ID,
)
from .context import ContextDocument, build_context
from .contract import (
    FailureCode,
    WriteResult,
    build_target_contract,
    make_prediction_record,
    make_trace_record,
    write_response_atomic,
)
from .prompts import (
    SYSTEM_PROMPT,
    assert_authored_budget,
    prompt_scaffold,
    target_summary_lines,
    workbook_char_budget,
)
from .provider import ANSWER_TOOL_NAME, CompletionProvider, ProviderConfigurationError


@dataclass(frozen=True, slots=True)
class TaskRun:
    task_id: str
    prediction: dict[str, object]
    success: bool


def _request_trace_metadata() -> dict[str, object]:
    """Describe the fixed request separately from provider-returned evidence."""

    return {
        "model": MODEL_ID,
        "model_provenance": MODEL_PROVENANCE,
        "max_output_tokens": DEFAULT_MAX_OUTPUT_TOKENS,
        "max_retries": PROVIDER_MAX_RETRIES,
        "num_samples": NUM_SAMPLES,
        "reasoning_effort": REASONING_EFFORT,
        "renderer": RENDERER_ID,
        "streaming": False,
        "stop": STOP_TOKEN,
        "temperature": TEMPERATURE,
        "telemetry": TINKER_TELEMETRY_MODE,
        "thinking": THINKING_MODE,
        "tool": ANSWER_TOOL_NAME,
        "tool_enforcement": "rendered_prompt_plus_strict_parser",
        "tokenizer_revision": TOKENIZER_REVISION,
        "transport": TRANSPORT_ID,
    }


def _safe_error(code: FailureCode, error: BaseException | None = None) -> str:
    """Return useful diagnostics without reflecting provider or dataset content."""

    suffix = f":{type(error).__name__}" if error is not None else ""
    if isinstance(error, ProviderConfigurationError) and error.status_code is not None:
        suffix += f":http_status={error.status_code}"
        if error.underlying_exception_class is not None:
            suffix += f":underlying_exception_class={error.underlying_exception_class}"
    return f"{code.value}{suffix}"


def _failure_write_result(
    *,
    task: DatasetTask,
    output_path: Path,
    code: FailureCode,
    error: BaseException | None = None,
) -> WriteResult:
    atomic_copy(task.init_xlsx, output_path)
    return WriteResult(
        output_path=output_path,
        success=False,
        failure_codes=(code,),
        error=_safe_error(code, error),
    )


def _build_prompt(task: DatasetTask) -> tuple[str, ContextDocument]:
    workbook = openpyxl.load_workbook(
        task.init_xlsx,
        read_only=False,
        data_only=False,
        keep_links=False,
    )
    try:
        contract = build_target_contract(task.as_dict(), workbook)
    finally:
        workbook.close()

    prefix, suffix = prompt_scaffold(
        instruction=task.instruction,
        target_lines=target_summary_lines(contract.ranges),
    )
    context = build_context(
        task.as_dict(),
        char_budget=workbook_char_budget(prefix=prefix, suffix=suffix),
    )
    prompt = prefix + context.jsonl + suffix
    assert_authored_budget(prompt)
    return prompt, context


async def run_task(
    *,
    task: DatasetTask,
    provider: CompletionProvider,
    layout: ArtifactLayout,
    semaphore: asyncio.Semaphore,
    staged_output_path: Path | None = None,
    staged_trace_path: Path | None = None,
) -> TaskRun:
    """Run one task and always attempt to publish its required artefacts."""

    relative_output = Path("outputs") / f"{task.id}.xlsx"
    output_path = staged_output_path or layout.root / relative_output
    trace_path = staged_trace_path or layout.traces / f"{task.id}.jsonl"
    prompt: str | None = None
    context: ContextDocument | None = None
    completion = None
    call_started = False
    started = time.perf_counter()

    try:
        prompt, context = _build_prompt(task)
    except Exception as exc:
        result = _failure_write_result(
            task=task,
            output_path=output_path,
            code=FailureCode.CONTEXT_BUILD_FAILED,
            error=exc,
        )
        atomic_write_jsonl(trace_path, [])
        prediction = make_prediction_record(task.id, relative_output, result)
        return TaskRun(task.id, prediction, False)

    try:
        async with semaphore:
            call_started = True
            completion = await provider.complete(
                system_prompt=SYSTEM_PROMPT,
                user_prompt=prompt,
            )
        write_result = write_response_atomic(
            task.as_dict(),
            completion.text,
            task.init_xlsx,
            output_path,
        )
        trace = make_trace_record(
            step=1,
            model=completion.model,
            prompt=prompt,
            response=completion.text,
            input_tokens=completion.input_tokens,
            output_tokens=completion.output_tokens,
            latency_ms=int((time.perf_counter() - started) * 1000),
            write_result=write_result,
            error=None if write_result.success else write_result.status,
        )
        trace["renderer"] = completion.renderer
        trace["stop_reason"] = completion.stop_reason
        trace["parse_termination"] = completion.parse_termination
        trace["response_format"] = completion.response_format
        trace["response_rejection"] = None
        trace["model_provenance"] = completion.model_provenance
        trace["request"] = _request_trace_metadata()
        trace["context"] = {
            "characters": context.used_chars,
            "included_cells": context.included_cells,
            "omitted_cells": context.omitted_cells,
            "truncated": context.truncated,
            "input_sha256": context.input_sha256,
        }
        atomic_write_jsonl(trace_path, [trace])
        prediction = make_prediction_record(task.id, relative_output, write_result)
        return TaskRun(task.id, prediction, write_result.success)
    except Exception as exc:
        write_result = _failure_write_result(
            task=task,
            output_path=output_path,
            code=FailureCode.MODEL_CALL_FAILED,
            error=exc,
        )
        if call_started:
            trace = make_trace_record(
                step=1,
                model=MODEL_ID,
                prompt=prompt,
                response=None,
                input_tokens=getattr(exc, "input_tokens", None),
                output_tokens=getattr(exc, "output_tokens", None),
                latency_ms=int((time.perf_counter() - started) * 1000),
                write_result=write_result,
                error=_safe_error(FailureCode.MODEL_CALL_FAILED, exc),
            )
            trace["renderer"] = RENDERER_ID
            trace["stop_reason"] = (
                completion.stop_reason
                if completion is not None
                else getattr(exc, "stop_reason", None)
            )
            trace["parse_termination"] = (
                completion.parse_termination
                if completion is not None
                else getattr(exc, "parse_termination", None)
            )
            trace["response_format"] = (
                completion.response_format if completion is not None else None
            )
            trace["response_rejection"] = (
                None if completion is not None else getattr(exc, "response_rejection", None)
            )
            trace["model_provenance"] = MODEL_PROVENANCE
            trace["request"] = _request_trace_metadata()
            atomic_write_jsonl(trace_path, [trace])
        else:
            atomic_write_jsonl(trace_path, [])
        prediction = make_prediction_record(task.id, relative_output, write_result)
        return TaskRun(task.id, prediction, False)


async def run_tasks(
    *,
    tasks: list[DatasetTask],
    provider: CompletionProvider,
    layout: ArtifactLayout,
    concurrency: int,
    existing_predictions: Mapping[str, Mapping[str, object]] | None = None,
    retry_task_ids: frozenset[str] | None = None,
) -> list[TaskRun]:
    """Run missing or explicitly retried tasks and checkpoint each completed result.

    ``tasks`` defines manifest order.  Validated existing predictions remain in
    the checkpoint map while work is in flight.  They are skipped unless their
    ID is explicitly present in ``retry_task_ids``.  After every completed task,
    its replacement and every still-pending old record are atomically rewritten
    in deterministic manifest order.
    """

    if concurrency < 1:
        raise ValueError("concurrency must be positive")
    task_ids = [task.id for task in tasks]
    if len(task_ids) != len(set(task_ids)):
        raise ValueError("task ids must be unique")

    existing_predictions = existing_predictions or {}
    retry_task_ids = retry_task_ids or frozenset()
    if unknown := set(existing_predictions).difference(task_ids):
        del unknown
        raise ValueError("existing prediction id is not in the selected tasks")
    if unknown := set(retry_task_ids).difference(task_ids):
        del unknown
        raise ValueError("retry task id is not in the selected tasks")
    if missing := set(retry_task_ids).difference(existing_predictions):
        del missing
        raise ValueError("retry task id has no existing prediction")
    for task_id in retry_task_ids:
        if existing_predictions[task_id].get("status") == "ok":
            raise ValueError("successful predictions cannot be retried")

    results_by_id: dict[str, TaskRun] = {}
    for task_id, source_record in existing_predictions.items():
        record = dict(source_record)
        if record.get("id") != task_id:
            raise ValueError("existing prediction key does not match its record id")
        results_by_id[task_id] = TaskRun(
            task_id=task_id,
            prediction=record,
            success=record.get("status") == "ok",
        )

    pending_ids = {
        task.id for task in tasks if task.id not in results_by_id or task.id in retry_task_ids
    }
    semaphore = asyncio.Semaphore(concurrency)
    task_by_id = {task.id: task for task in tasks}
    retry_workspaces: dict[str, RetryWorkspace] = {}
    try:
        for task in tasks:
            if task.id in retry_task_ids:
                retry_workspaces[task.id] = create_retry_workspace(layout, task.id)
    except BaseException:
        for workspace in retry_workspaces.values():
            with suppress(BaseException):
                discard_retry_workspace(layout, workspace)
        raise

    workers = []
    for task in tasks:
        if task.id not in pending_ids:
            continue
        workspace = retry_workspaces.get(task.id)
        workers.append(
            asyncio.create_task(
                run_task(
                    task=task,
                    provider=provider,
                    layout=layout,
                    semaphore=semaphore,
                    staged_output_path=workspace.new_output if workspace is not None else None,
                    staged_trace_path=workspace.new_trace if workspace is not None else None,
                )
            )
        )
    checkpointed_worker_ids: set[str] = set()
    commit_started_ids: set[str] = set()

    def checkpoint() -> None:
        ordered = [
            results_by_id[task_id].prediction for task_id in task_ids if task_id in results_by_id
        ]
        atomic_write_jsonl(layout.predictions, ordered)

    def publish_result(result: TaskRun) -> None:
        if result.task_id in retry_task_ids:
            commit_started_ids.add(result.task_id)
            workspace = retry_workspaces[result.task_id]
            prepare_retry_transaction(
                workspace,
                task=task_by_id[result.task_id],
                old_prediction=results_by_id[result.task_id].prediction,
                new_prediction=result.prediction,
                expected_model=MODEL_ID,
            )
            committed = commit_retry_transaction(
                layout,
                tasks,
                task_id=result.task_id,
                expected_model=MODEL_ID,
            )
            if committed != result.prediction:
                raise RuntimeError("retry commit returned an unexpected prediction")
            results_by_id[result.task_id] = result
        else:
            results_by_id[result.task_id] = result
            checkpoint()
        checkpointed_worker_ids.add(result.task_id)

    main_loop_completed = False
    observed_workers: set[asyncio.Task[TaskRun]] = set()
    deferred_errors: list[BaseException] = []

    def drain_finished_workers() -> None:
        for worker in workers:
            if worker in observed_workers or not worker.done():
                continue
            observed_workers.add(worker)
            if worker.cancelled():
                continue
            try:
                result = worker.result()
                if (
                    result.task_id not in checkpointed_worker_ids
                    and result.task_id not in commit_started_ids
                ):
                    publish_result(result)
            except BaseException as exc:
                deferred_errors.append(exc)

    try:
        for completed in asyncio.as_completed(workers):
            result = await completed
            publish_result(result)
        main_loop_completed = True
    finally:
        # Cancellation can arrive after a worker has published its workbook and
        # trace but just before as_completed yields it.  Drain those finished
        # results into one final atomic checkpoint before cancelling the rest.
        drain_finished_workers()
        unfinished = [worker for worker in workers if not worker.done()]
        for worker in unfinished:
            worker.cancel()
        if unfinished:
            await asyncio.gather(*unfinished, return_exceptions=True)
        drain_finished_workers()
        for workspace in retry_workspaces.values():
            if workspace.directory.exists() and not (workspace.directory / "meta.json").exists():
                try:
                    discard_retry_workspace(layout, workspace)
                except BaseException as exc:
                    deferred_errors.append(exc)
        if main_loop_completed and deferred_errors:
            raise deferred_errors[0]

    return [results_by_id[task_id] for task_id in task_ids]
