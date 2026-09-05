"""Production FormulaBench command line entry point."""

from __future__ import annotations

import argparse
import asyncio
import os
import sys
from collections.abc import Sequence
from pathlib import Path

from .artifacts import (
    ArtifactError,
    ArtifactLayout,
    DatasetTask,
    ResumeStateError,
    exclusive_output_lock,
    load_dataset_manifest,
    load_resume_state,
    recover_retry_transactions,
)
from .constants import KEY_ENV_VAR, MODEL_ID, RENDERER_ID, THINKING_MODE, TRANSPORT_ID
from .provider import (
    ANSWER_TOOL_NAME,
    ProviderConfigurationError,
    TinkerProvider,
    require_provider_key,
    validate_native_runtime,
)
from .replay import replay_stored_responses
from .runner import TaskRun, run_tasks
from .validate_out import validate_output


def _positive_concurrency(value: str) -> int:
    try:
        parsed = int(value)
    except ValueError as exc:
        raise argparse.ArgumentTypeError("concurrency must be an integer") from exc
    if not 1 <= parsed <= 64:
        raise argparse.ArgumentTypeError("concurrency must be between 1 and 64")
    return parsed


def parse_args(argv: Sequence[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Run the fixed FormulaBench inference pipeline.")
    parser.add_argument("--dataset-dir", required=True)
    parser.add_argument("--out-dir", required=True)
    parser.add_argument("--ids", help="comma-separated task ids for a local partial run")
    parser.add_argument("--concurrency", type=_positive_concurrency, default=4)
    parser.add_argument("--preflight-only", action="store_true")
    parser.add_argument(
        "--resume",
        action="store_true",
        help="validate and continue an interrupted or partial output directory",
    )
    parser.add_argument(
        "--retry-failures",
        action="store_true",
        help="with --resume, deliberately rerun validated failed checkpoints",
    )
    parser.add_argument(
        "--replay-write-failures",
        action="store_true",
        help="with --resume, rematerialise stored workbook-write responses without a model call",
    )
    args = parser.parse_args(argv)
    if args.retry_failures and not args.resume:
        parser.error("--retry-failures requires --resume")
    if args.replay_write_failures and not args.resume:
        parser.error("--replay-write-failures requires --resume")
    if args.retry_failures and args.replay_write_failures:
        parser.error("--retry-failures and --replay-write-failures are mutually exclusive")
    return args


def _select_tasks(tasks: list[DatasetTask], ids: str | None) -> list[DatasetTask]:
    if ids is None:
        return tasks
    requested = [item.strip() for item in ids.split(",") if item.strip()]
    requested_set = set(requested)
    if len(requested) != len(requested_set):
        raise ArtifactError("--ids contains duplicate task ids")
    known = {task.id for task in tasks}
    if unknown := requested_set - known:
        raise ArtifactError(f"--ids contains {len(unknown)} unknown task id(s)")
    return [task for task in tasks if task.id in requested_set]


def _ensure_output_outside_dataset(dataset_dir: Path, out_dir: Path) -> None:
    if out_dir == dataset_dir or dataset_dir in out_dir.parents:
        raise ArtifactError("output directory must be outside the dataset")


def _preflight(tasks: list[DatasetTask], selected: list[DatasetTask]) -> None:
    validate_native_runtime()
    print(
        f"preflight ok  tasks={len(tasks)}  selected={len(selected)}  "
        f"model={MODEL_ID}  transport={TRANSPORT_ID}  "
        f"renderer={RENDERER_ID}  thinking={THINKING_MODE}  "
        f"tool_enforcement=rendered_prompt_plus_strict_parser:{ANSWER_TOOL_NAME}",
        flush=True,
    )


async def _run(args: argparse.Namespace) -> int:
    dataset_dir = Path(args.dataset_dir).resolve(strict=True)
    out_dir = Path(args.out_dir).resolve(strict=False)
    _ensure_output_outside_dataset(dataset_dir, out_dir)
    tasks = load_dataset_manifest(dataset_dir)
    selected = _select_tasks(tasks, args.ids)
    if not selected:
        raise ArtifactError("no tasks selected")

    if args.preflight_only:
        _preflight(tasks, selected)
        return 0

    layout = (
        ArtifactLayout.resume(out_dir)
        if args.resume
        else ArtifactLayout.initialise(out_dir, allow_supervisor_log=True)
    )
    with exclusive_output_lock(layout.run_log):
        return await _run_owned(args, dataset_dir, out_dir, tasks, selected, layout)


async def _run_owned(
    args: argparse.Namespace,
    dataset_dir: Path,
    out_dir: Path,
    tasks: list[DatasetTask],
    selected: list[DatasetTask],
    layout: ArtifactLayout,
) -> int:
    """Run while the caller holds exclusive ownership of the output root."""

    retry_task_ids: frozenset[str] = frozenset()
    if args.resume:
        recover_retry_transactions(layout, selected, expected_model=MODEL_ID)
        resume_state = load_resume_state(layout, selected, expected_model=MODEL_ID)
        if args.replay_write_failures:
            if len(resume_state.completed_ids) != len(selected):
                raise ResumeStateError("replay_requires_complete_run")
            replay = replay_stored_responses(
                layout=layout,
                tasks=selected,
                expected_model=MODEL_ID,
            )
            print(
                f"stored-response replay  eligible={replay.eligible_tasks}  "
                f"replayed={len(replay.replayed_tasks)}  "
                f"unchanged={len(replay.still_failed)}  "
                f"additional_model_calls={replay.model_calls}",
                flush=True,
            )
            resume_state = load_resume_state(layout, selected, expected_model=MODEL_ID)
        existing_predictions = resume_state.predictions
        if args.retry_failures:
            retry_task_ids = resume_state.completed_ids.difference(resume_state.successful_ids)
        print(
            f"resume validated  completed={len(resume_state.completed_ids)}  "
            f"missing={len(selected) - len(resume_state.completed_ids)}  "
            f"retrying_failures={len(retry_task_ids)}",
            flush=True,
        )
    else:
        existing_predictions = {}

    pending_count = len(selected) - len(existing_predictions) + len(retry_task_ids)
    if pending_count:
        require_provider_key()
        provider = TinkerProvider()
        try:
            print(
                f"FormulaBench  model={MODEL_ID}  tasks={len(selected)}  "
                f"remaining={pending_count}  concurrency={args.concurrency}",
                flush=True,
            )
            results = await run_tasks(
                tasks=selected,
                provider=provider,
                layout=layout,
                concurrency=args.concurrency,
                existing_predictions=existing_predictions,
                retry_task_ids=retry_task_ids,
            )
        finally:
            await provider.close()
    else:
        print(
            f"FormulaBench  model={MODEL_ID}  tasks={len(selected)}  remaining=0",
            flush=True,
        )
        results = [
            TaskRun(
                task_id=task.id,
                prediction=dict(existing_predictions[task.id]),
                success=existing_predictions[task.id].get("status") == "ok",
            )
            for task in selected
        ]

    completed_state = load_resume_state(layout, selected, expected_model=MODEL_ID)
    if len(completed_state.completed_ids) != len(selected):
        raise ResumeStateError("run_incomplete")
    successes = sum(result.success for result in results)
    print(f"completed  ok={successes}  failed={len(results) - successes}", flush=True)

    if len(selected) != len(tasks):
        print("partial run: full output validation deferred", flush=True)
        return 1 if successes == 0 else 0

    secret = os.environ.get(KEY_ENV_VAR, "")
    report = validate_output(
        dataset_dir,
        out_dir,
        expected_model=MODEL_ID,
        secret_values=(secret,),
        require_traces=True,
    )
    print(
        f"output validation  ok={report.ok}  issues={len(report.issues)}  "
        f"workbooks={report.readable_workbooks}/{report.expected_tasks}",
        flush=True,
    )
    if not report.ok:
        codes = sorted({issue.code for issue in report.issues})
        print(f"output validation codes: {','.join(codes)}", flush=True)
        return 1
    if successes == 0:
        print("run failed: no task produced a generated answer", file=sys.stderr, flush=True)
        return 1
    return 0


def main(argv: Sequence[str] | None = None) -> int:
    args = parse_args(argv)
    try:
        return asyncio.run(_run(args))
    except ResumeStateError as exc:
        print(f"FormulaBench failed: {exc}", file=sys.stderr, flush=True)
        return 2
    except (ArtifactError, ProviderConfigurationError, OSError, ValueError) as exc:
        print(f"FormulaBench failed: {type(exc).__name__}", file=sys.stderr, flush=True)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
