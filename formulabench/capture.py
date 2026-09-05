"""Run FormulaBench while preserving the child's raw combined output."""

from __future__ import annotations

import argparse
import os
import signal
import stat
import subprocess
import sys
from collections.abc import Mapping, Sequence
from dataclasses import dataclass
from pathlib import Path

from .artifacts import (
    SUPERVISOR_LOCK_FD_ENV,
    ArtifactError,
    ResumeStateError,
    load_dataset_manifest,
    normalise_task_id,
    require_resume_directory,
    require_writable_directory,
)

try:  # The supported host and container are POSIX.
    import fcntl
except ImportError:  # pragma: no cover - importability on non-POSIX systems.
    fcntl = None


@dataclass(frozen=True, slots=True)
class CaptureResult:
    command: tuple[str, ...]
    returncode: int
    run_log: Path


def _forward_signal(process: subprocess.Popen[bytes], signum: int) -> None:
    if process.poll() is not None:
        return
    try:
        if os.name == "posix":
            os.killpg(process.pid, signum)
        else:  # pragma: no cover - Docker and the supported host are POSIX.
            process.send_signal(signum)
    except ProcessLookupError:
        pass


def _echo_raw(chunk: bytes) -> None:
    """Best-effort live echo; the authoritative bytes are already in run.log."""
    try:
        stream = getattr(sys.stdout, "buffer", None)
        if stream is not None:
            stream.write(chunk)
            stream.flush()
        else:  # Useful under test runners that replace sys.stdout.
            sys.stdout.write(chunk.decode("utf-8", errors="replace"))
            sys.stdout.flush()
    except (BrokenPipeError, OSError):
        pass


def run_captured(
    command: Sequence[str],
    out_dir: str | os.PathLike[str],
    *,
    cwd: str | os.PathLike[str] | None = None,
    env: Mapping[str, str] | None = None,
    echo: bool = True,
    resume: bool = False,
) -> CaptureResult:
    """Run ``command`` with stderr merged into stdout and copied to run.log.

    A new run requires an empty writable output directory.  A resumed run
    requires an existing canonical output root and opens ``run.log`` with
    append-only semantics; all existing bytes are preserved.  Nothing is
    cleaned up or overwritten.  Signals received by the supervisor are
    forwarded to the child's process group so descendants do not remain.
    """
    if not command or any(not isinstance(part, str) or not part for part in command):
        raise ArtifactError("capture command must contain non-empty strings")
    root = require_resume_directory(out_dir) if resume else require_writable_directory(out_dir)
    run_log = root / "run.log"

    # Exclusive creation protects a new run; O_APPEND protects every byte from
    # earlier attempts during resume.  O_NOFOLLOW closes the symlink race after
    # the directory check.  The log is never rewritten or redacted.
    flags = os.O_WRONLY | getattr(os, "O_CLOEXEC", 0) | getattr(os, "O_NOFOLLOW", 0)
    flags |= os.O_APPEND if resume else os.O_CREAT | os.O_EXCL
    log_fd = os.open(run_log, flags, 0o600)
    if not stat.S_ISREG(os.fstat(log_fd).st_mode):
        os.close(log_fd)
        raise (
            ResumeStateError("run_log_unsafe")
            if resume
            else ArtifactError("run.log is not a regular file")
        )
    if fcntl is not None:
        try:
            fcntl.flock(log_fd, fcntl.LOCK_EX | fcntl.LOCK_NB)
        except BlockingIOError as exc:
            os.close(log_fd)
            if resume:
                raise ResumeStateError("run_already_active") from exc
            raise ArtifactError("another run is already active") from exc
    process: subprocess.Popen[bytes] | None = None
    previous_handlers: dict[int, signal.Handlers] = {}
    pending_signals: list[int] = []
    handled_signals = [signal.SIGINT, signal.SIGTERM]
    for optional_name in ("SIGHUP", "SIGQUIT"):
        optional_signal = getattr(signal, optional_name, None)
        if optional_signal is not None:
            handled_signals.append(optional_signal)

    def handler(signum: int, _frame: object) -> None:
        if process is None:
            pending_signals.append(signum)
        else:
            _forward_signal(process, signum)

    try:
        # Install handlers before spawning so there is no window in which the
        # supervisor can exit and orphan a newly created child process group.
        try:
            for signum in handled_signals:
                previous = signal.getsignal(signum)
                signal.signal(signum, handler)
                previous_handlers[signum] = previous
        except ValueError:
            # Signal handlers may only be installed in the main thread.  The
            # production CLI always runs there; library callers still capture.
            for signum, previous in previous_handlers.items():
                signal.signal(signum, previous)
            previous_handlers.clear()

        child_env = dict(env) if env is not None else dict(os.environ)
        popen_options: dict[str, object] = {}
        if os.name == "posix" and fcntl is not None:
            child_env[SUPERVISOR_LOCK_FD_ENV] = str(log_fd)
            popen_options["pass_fds"] = (log_fd,)
        process = subprocess.Popen(
            list(command),
            cwd=os.fspath(cwd) if cwd is not None else None,
            env=child_env,
            stdin=None,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            bufsize=0,
            start_new_session=os.name == "posix",
            **popen_options,
        )
        for pending_signal in pending_signals:
            _forward_signal(process, pending_signal)

        assert process.stdout is not None
        while chunk := process.stdout.read(64 * 1024):
            view = memoryview(chunk)
            while view:
                written = os.write(log_fd, view)
                if written <= 0:
                    raise OSError("short write to run.log")
                view = view[written:]
            if echo:
                _echo_raw(chunk)
        returncode = process.wait()
        os.fsync(log_fd)
    except BaseException:
        if process is not None and process.poll() is None:
            _forward_signal(process, signal.SIGTERM)
            try:
                process.wait(timeout=5)
            except subprocess.TimeoutExpired:
                _forward_signal(process, signal.SIGKILL)
                process.wait()
        raise
    finally:
        for signum, previous in previous_handlers.items():
            signal.signal(signum, previous)
        os.close(log_fd)

    return CaptureResult(tuple(command), returncode, run_log)


def _positive_concurrency(value: str) -> int:
    try:
        parsed = int(value)
    except ValueError as exc:
        raise argparse.ArgumentTypeError("concurrency must be an integer") from exc
    if not 1 <= parsed <= 64:
        raise argparse.ArgumentTypeError("concurrency must be between 1 and 64")
    return parsed


def _validated_ids(value: str | None, dataset_ids: set[str]) -> str | None:
    if value is None:
        return None
    raw_ids = [item.strip() for item in value.split(",")]
    if not raw_ids or any(not item for item in raw_ids):
        raise ArtifactError("--ids must be a non-empty comma-separated list")
    task_ids: list[str] = []
    seen: set[str] = set()
    for raw_id in raw_ids:
        task_id = normalise_task_id(raw_id)
        if task_id in seen:
            raise ArtifactError("--ids contains a duplicate task id")
        if task_id not in dataset_ids:
            raise ArtifactError("--ids contains a task not present in the dataset")
        seen.add(task_id)
        task_ids.append(task_id)
    return ",".join(task_ids)


def parse_args(argv: Sequence[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Run FormulaBench and capture unedited stdout/stderr in run.log."
    )
    parser.add_argument("--dataset-dir", required=True)
    parser.add_argument("--out-dir", required=True)
    parser.add_argument("--ids", help="comma-separated task ids")
    parser.add_argument("--concurrency", type=_positive_concurrency, default=4)
    parser.add_argument("--preflight-only", action="store_true")
    parser.add_argument(
        "--resume",
        action="store_true",
        help="continue a validated output directory and append to its run.log",
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


def main(argv: Sequence[str] | None = None) -> int:
    args = parse_args(argv)
    try:
        tasks = load_dataset_manifest(args.dataset_dir)
        dataset_root = Path(args.dataset_dir).resolve(strict=True)
        output_candidate = Path(args.out_dir).resolve(strict=False)
        if output_candidate == dataset_root or dataset_root in output_candidate.parents:
            raise ArtifactError("output directory must be outside the dataset")
        selected_ids = _validated_ids(args.ids, {task.id for task in tasks})

        command = [
            sys.executable,
            "-m",
            "formulabench.cli",
            f"--dataset-dir={dataset_root}",
            f"--out-dir={output_candidate}",
            f"--concurrency={args.concurrency}",
        ]
        if selected_ids is not None:
            command.append(f"--ids={selected_ids}")
        if args.preflight_only:
            command.append("--preflight-only")
        if args.resume:
            command.append("--resume")
        if args.retry_failures:
            command.append("--retry-failures")
        if args.replay_write_failures:
            command.append("--replay-write-failures")
        result = run_captured(command, output_candidate, resume=args.resume)
    except ResumeStateError as exc:
        print(f"capture failed: {exc}", file=sys.stderr)
        return 2
    except (ArtifactError, OSError, subprocess.SubprocessError) as exc:
        # Error text is intentionally generic: dataset contents and environment
        # values must not be reflected into logs or CI output.
        print(f"capture failed: {type(exc).__name__}", file=sys.stderr)
        return 2
    return result.returncode if result.returncode >= 0 else 128 - result.returncode


if __name__ == "__main__":
    raise SystemExit(main())
