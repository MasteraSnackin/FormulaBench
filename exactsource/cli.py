"""Command-line entry point for the `/data` to `/out` judge runtime."""

from __future__ import annotations

import argparse
import sys
import threading
from collections.abc import Callable, Sequence
from contextlib import contextmanager
from pathlib import Path
from typing import TextIO

from exactsource import __version__
from exactsource.artifacts import ArtifactError, OutputLayout, atomic_write_text, prepare_output
from exactsource.config import CONCURRENCY, MODEL_NAME, RuntimePaths
from exactsource.contracts import TaskSpec
from exactsource.dataset import load_tasks
from exactsource.runner import DefaultTaskSolver, TaskSolver, _safe_error, run_tasks

TaskLoader = Callable[[Path], list[TaskSpec]]


class _Tee:
    """Write the application's stream verbatim to the terminal and run log."""

    def __init__(self, terminal: TextIO, log: TextIO, lock: threading.Lock) -> None:
        self.terminal = terminal
        self.log = log
        self.lock = lock
        self.encoding = getattr(terminal, "encoding", "utf-8")

    def write(self, text: str) -> int:
        with self.lock:
            terminal_count = self.terminal.write(text)
            self.terminal.flush()
            self.log.write(text)
            self.log.flush()
        return terminal_count

    def flush(self) -> None:
        with self.lock:
            self.terminal.flush()
            self.log.flush()

    def isatty(self) -> bool:
        return False


@contextmanager
def _capture_run_log(layout: OutputLayout):
    atomic_write_text(layout.log_path, "")
    original_stdout = sys.stdout
    original_stderr = sys.stderr
    lock = threading.Lock()
    with layout.log_path.open("a", encoding="utf-8", newline="\n") as log:
        sys.stdout = _Tee(original_stdout, log, lock)  # type: ignore[assignment]
        sys.stderr = _Tee(original_stderr, log, lock)  # type: ignore[assignment]
        try:
            yield
        finally:
            sys.stdout = original_stdout
            sys.stderr = original_stderr


def _parser() -> argparse.ArgumentParser:
    defaults = RuntimePaths()
    parser = argparse.ArgumentParser(
        prog="exactsource",
        description="Run ExactSource against a SpreadsheetBench-format dataset.",
    )
    parser.add_argument("--data-dir", type=Path, default=defaults.data_dir)
    parser.add_argument("--out-dir", type=Path, default=defaults.out_dir)
    parser.add_argument(
        "--ids",
        help="comma-separated task ids for a development smoke run; default is every task",
    )
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    return parser


def _validated_mounts(data_dir: Path, out_dir: Path) -> tuple[Path, Path]:
    try:
        data = data_dir.resolve(strict=True)
    except (FileNotFoundError, OSError) as error:
        raise ArtifactError(f"dataset directory does not exist: {data_dir}") from error
    if not data.is_dir():
        raise ArtifactError(f"dataset path is not a directory: {data_dir}")
    output = out_dir.resolve(strict=False)
    if output == data or output in data.parents or data in output.parents:
        raise ArtifactError("dataset and output directories must not overlap")
    return data, output


def run_cli(
    argv: Sequence[str] | None = None,
    *,
    task_loader: TaskLoader = load_tasks,
    solver: TaskSolver | None = None,
) -> int:
    args = _parser().parse_args(argv)
    try:
        data_dir, out_dir = _validated_mounts(args.data_dir, args.out_dir)
        layout = prepare_output(out_dir)
    except Exception as error:
        print(f"ExactSource startup failed: {_safe_error(error)}", file=sys.stderr)
        return 2

    created_solver: DefaultTaskSolver | None = None
    with _capture_run_log(layout):
        print(f"ExactSource {__version__} model={MODEL_NAME} data={data_dir} out={out_dir}")
        try:
            tasks = task_loader(data_dir)
            if not tasks:
                raise ArtifactError("dataset contains no tasks")
            if args.ids:
                requested = {task_id.strip() for task_id in args.ids.split(",") if task_id.strip()}
                if not requested:
                    raise ArtifactError("--ids must contain at least one task id")
                available = {task.id for task in tasks}
                unknown = sorted(requested - available)
                if unknown:
                    raise ArtifactError(f"unknown task ids requested: {unknown[:5]}")
                tasks = [task for task in tasks if task.id in requested]
            active_solver = solver
            if active_solver is None:
                created_solver = DefaultTaskSolver()
                active_solver = created_solver
            summary = run_tasks(
                tasks,
                out_dir,
                active_solver,
                concurrency=CONCURRENCY,
            )
            print(
                f"ExactSource finished: {summary.succeeded}/{summary.total} tasks succeeded; "
                f"{summary.failed} used the safe fallback."
            )
            return 0
        except Exception as error:
            print(f"ExactSource run failed: {_safe_error(error)}", file=sys.stderr)
            return 1
        finally:
            if created_solver is not None:
                created_solver.close()


def main() -> int:
    return run_cli()


if __name__ == "__main__":
    raise SystemExit(main())
