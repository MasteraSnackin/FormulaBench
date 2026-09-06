"""FormulaBench v2 adapter for the vendored ExactSource engine.

The v2 path deliberately has a smaller surface than the original FormulaBench
runner.  It starts only fresh ExactSource runs; the legacy capture command is
still available behind the explicit ``--legacy-engine`` escape hatch.
"""

from __future__ import annotations

import argparse
import importlib.util
import os
import stat
import sys
from collections.abc import Iterator, Sequence
from contextlib import contextmanager, suppress
from functools import wraps
from pathlib import Path
from typing import Any

try:  # The submitted runtime is POSIX; keep import-time help usable elsewhere.
    import fcntl
except ImportError:  # pragma: no cover - exercised only on non-POSIX platforms.
    fcntl = None

V2_CONCURRENCY = 4
VENDORED_SOURCE_ID = "ExactSource@99fe8084bf35a5fca6a2c2e1c9beae802766a618"
SCORED_CORE_ID = "ExactSource@8b84dba1d9263e2123b8f15267239b70ff817907"
_API_KEY_ENV = "TINKER_API_KEY"
_PRIVATE_DIRECTORY_MODE = 0o700
_PRIVATE_FILE_MODE = 0o600
_CLEAN_ENV_REMEDY = (
    "use the Docker runner or a clean `uv sync --locked` environment without "
    "the `native-tinker` extra"
)


class V2Error(RuntimeError):
    """Raised when the v2 adapter cannot start a safe ExactSource run."""


def _configure_v2_serializer() -> None:
    """Match the scored ExactSource openpyxl serializer before importing it."""

    os.environ["OPENPYXL_LXML"] = "False"


def _v2_concurrency(value: str) -> int:
    try:
        concurrency = int(value)
    except ValueError as exc:
        raise argparse.ArgumentTypeError("concurrency must be an integer") from exc
    if concurrency != V2_CONCURRENCY:
        raise argparse.ArgumentTypeError(
            f"FormulaBench v2 requires concurrency {V2_CONCURRENCY} for ExactSource parity"
        )
    return concurrency


def _engine_version() -> str:
    from exactsource import __version__

    return __version__


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="python -m formulabench.v2",
        description="Run FormulaBench with the vendored ExactSource engine.",
        allow_abbrev=False,
    )
    parser.add_argument("--dataset-dir", required=True)
    parser.add_argument("--out-dir", required=True)
    parser.add_argument("--ids", help="comma-separated task ids")
    parser.add_argument(
        "--concurrency",
        type=_v2_concurrency,
        default=V2_CONCURRENCY,
        help="fixed at 4 for ExactSource parity (default: 4)",
    )
    parser.add_argument("--preflight-only", action="store_true")
    parser.add_argument(
        "--legacy-engine",
        action="store_true",
        help="run the original FormulaBench engine instead of v2",
    )
    parser.add_argument(
        "--version",
        action="version",
        version=(
            f"%(prog)s ExactSource/{_engine_version()} "
            f"vendored_source={VENDORED_SOURCE_ID} scored_core={SCORED_CORE_ID}"
        ),
    )

    # Parse these legacy-only options so v2 can explain the safe alternative
    # precisely instead of reporting a generic "unrecognised argument" error.
    parser.add_argument("--resume", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--retry-failures", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--replay-write-failures", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--sampler-checkpoint", help=argparse.SUPPRESS)
    return parser


def parse_args(argv: Sequence[str] | None = None) -> argparse.Namespace:
    # _parser() obtains the vendored engine version, so the serializer setting
    # must precede even parser construction.  main() handles the legacy escape
    # hatch before reaching this v2-only function.
    _configure_v2_serializer()
    parser = _parser()
    args = parser.parse_args(argv)
    if args.sampler_checkpoint is not None:
        parser.error(
            "--sampler-checkpoint is not supported by FormulaBench v2 or its "
            "--legacy-engine adapter; use `python -m formulabench.cli` directly "
            "for checkpoint experiments"
        )
    unsupported = (
        ("--resume", args.resume),
        ("--retry-failures", args.retry_failures),
        ("--replay-write-failures", args.replay_write_failures),
    )
    for option, present in unsupported:
        if present:
            parser.error(
                f"{option} is not supported by FormulaBench v2; "
                "use --legacy-engine to run the legacy engine"
            )
    return args


def _validated_paths(dataset_dir: str, out_dir: str) -> tuple[Path, Path]:
    dataset_candidate = Path(dataset_dir)
    try:
        dataset = dataset_candidate.resolve(strict=True)
    except (FileNotFoundError, OSError, RuntimeError) as exc:
        raise V2Error(f"dataset directory does not exist: {dataset_candidate}") from exc
    if not dataset.is_dir():
        raise V2Error(f"dataset path is not a directory: {dataset_candidate}")

    output_candidate = Path(out_dir)
    if output_candidate.is_symlink():
        raise V2Error("output directory must not be a symbolic link")
    try:
        output = output_candidate.resolve(strict=False)
    except (OSError, RuntimeError) as exc:
        raise V2Error("output directory cannot be resolved") from exc
    if output == dataset or output in dataset.parents or dataset in output.parents:
        raise V2Error("dataset and output directories must not overlap")
    return dataset, output


def _load_tasks(dataset_dir: Path) -> list[Any]:
    # Kept behind a narrow wrapper so importing this module (and asking for
    # --help/--version) never initialises a provider client.
    from exactsource.dataset import load_tasks

    return load_tasks(dataset_dir)


def _select_tasks(tasks: list[Any], ids: str | None) -> list[Any]:
    """Apply ExactSource's development-id selection semantics.

    Selection is set-based while returned tasks retain manifest order, matching
    the vendored CLI.  Repeated requested ids therefore do not duplicate work.
    """

    if ids is None:
        return tasks
    requested = {task_id.strip() for task_id in ids.split(",") if task_id.strip()}
    if not requested:
        raise V2Error("--ids must contain at least one task id")
    available = {task.id for task in tasks}
    unknown = sorted(requested - available)
    if unknown:
        raise V2Error(f"unknown task ids requested: {unknown[:5]}")
    return [task for task in tasks if task.id in requested]


def _preflight(tasks: list[Any], selected: list[Any]) -> None:
    from exactsource.config import MODEL_NAME

    print(
        f"preflight ok  tasks={len(tasks)}  selected={len(selected)}  "
        f"engine=ExactSource/{_engine_version()}  model={MODEL_NAME}  "
        f"vendored_source={VENDORED_SOURCE_ID}  scored_core={SCORED_CORE_ID}",
        flush=True,
    )


def _pillow_available() -> bool:
    """Return whether the non-canonical Pillow serializer dependency is present."""

    return importlib.util.find_spec("PIL") is not None


def _openpyxl_uses_lxml() -> bool:
    """Inspect openpyxl's effective serializer, including a preloaded module."""

    from openpyxl.xml import functions as xml_functions

    return xml_functions.LXML is not False


def _require_canonical_runtime() -> None:
    """Fail before inference if workbook bytes could differ from the scored run."""

    if _pillow_available():
        raise V2Error(
            "Pillow/PIL is installed, but FormulaBench v2 requires the canonical "
            f"Pillow-free dependency surface; {_CLEAN_ENV_REMEDY}"
        )
    if _openpyxl_uses_lxml():
        raise V2Error(
            "openpyxl is using its lxml serializer, but FormulaBench v2 requires "
            f"the stdlib XML serializer; {_CLEAN_ENV_REMEDY}"
        )


def _require_api_key() -> None:
    """Reject a paid run before output creation when its credential is absent."""

    value = os.environ.get(_API_KEY_ENV)
    if value is None or not value.strip():
        raise V2Error(
            f"{_API_KEY_ENV} is not set; supply it privately through the process environment"
        )


def _run_exactsource(argv: Sequence[str]) -> int:
    _configure_v2_serializer()
    _require_canonical_runtime()
    _install_exactsource_private_permissions()
    _install_exactsource_child_serializer()
    from exactsource.cli import run_cli

    return run_cli(argv)


def _install_exactsource_child_serializer() -> None:
    """Add the scored serializer setting to ExactSource's clean child env.

    ExactSource intentionally constructs an allow-listed environment rather
    than inheriting the parent process.  Wrapping that existing builder keeps
    the security boundary intact and changes only the XML serializer selection.
    The marker makes repeated in-process v2 invocations idempotent.
    """

    from exactsource import sandbox

    child_environment = sandbox._child_environment
    marker = "__formulabench_v2_openpyxl_lxml__"
    if getattr(child_environment, marker, False):
        return

    @wraps(child_environment)
    def v2_child_environment(timeout: float) -> dict[str, str]:
        environment = dict(child_environment(timeout))
        environment["OPENPYXL_LXML"] = "False"
        return environment

    setattr(v2_child_environment, marker, True)
    sandbox._child_environment = v2_child_environment


def _private_plan_save(original: Any) -> Any:
    """Wrap ExactSource's operations-route save with v2's private file mode."""

    @wraps(original)
    def save(workbook: object, destination: Path) -> None:
        try:
            original(workbook, destination)
        finally:
            _harden_existing_regular_file(Path(destination))

    marker = "__formulabench_v2_private_permissions__"
    setattr(save, marker, True)
    return save


def _private_sandbox_promote(original: Any) -> Any:
    """Wrap ExactSource's Python-route promotion with v2's private file mode."""

    @wraps(original)
    def promote(source: Path, destination: Path) -> None:
        try:
            original(source, destination)
        finally:
            _harden_existing_regular_file(Path(destination))

    marker = "__formulabench_v2_private_permissions__"
    setattr(promote, marker, True)
    return promote


def _install_exactsource_private_permissions() -> None:
    """Adapt vendored ExactSource's public artefact modes to private v2 modes."""

    from exactsource import artifacts, plans, sandbox

    artifacts._PUBLISHED_DIRECTORY_MODE = _PRIVATE_DIRECTORY_MODE
    artifacts._PUBLISHED_FILE_MODE = _PRIVATE_FILE_MODE

    marker = "__formulabench_v2_private_permissions__"
    if not getattr(plans._save_atomic, marker, False):
        plans._save_atomic = _private_plan_save(plans._save_atomic)
    if not getattr(sandbox._promote_atomic, marker, False):
        sandbox._promote_atomic = _private_sandbox_promote(sandbox._promote_atomic)


def _harden_existing_regular_file(path: Path) -> None:
    """Set a present regular file private without following a symbolic link."""

    try:
        path_stat = path.lstat()
    except FileNotFoundError:
        return
    if not stat.S_ISREG(path_stat.st_mode):
        raise V2Error("v2 output contains an unsafe non-regular file")
    _fchmod_path(path, expected_directory=False, mode=_PRIVATE_FILE_MODE)


def _fchmod_path(path: Path, *, expected_directory: bool, mode: int) -> None:
    """Open, type-check and chmod one path without following a leaf symlink."""

    flags = os.O_RDONLY | getattr(os, "O_CLOEXEC", 0) | getattr(os, "O_NOFOLLOW", 0)
    if expected_directory:
        flags |= getattr(os, "O_DIRECTORY", 0)
    descriptor = os.open(path, flags)
    try:
        descriptor_stat = os.fstat(descriptor)
        expected = stat.S_ISDIR if expected_directory else stat.S_ISREG
        if not expected(descriptor_stat.st_mode):
            raise V2Error("v2 output contains an unsafe filesystem entry")
        os.fchmod(descriptor, mode)
    finally:
        os.close(descriptor)


def _harden_output_permissions(out_dir: Path) -> None:
    """Make every v2 directory and regular artefact private, including failures."""

    if out_dir.is_symlink() or not out_dir.is_dir():
        raise V2Error("v2 output root became unsafe during the run")
    for current, directory_names, file_names in os.walk(out_dir, followlinks=False):
        current_path = Path(current)
        _fchmod_path(
            current_path,
            expected_directory=True,
            mode=_PRIVATE_DIRECTORY_MODE,
        )
        for name in directory_names:
            child = current_path / name
            if child.is_symlink():
                raise V2Error("v2 output contains a symbolic link")
            _fchmod_path(
                child,
                expected_directory=True,
                mode=_PRIVATE_DIRECTORY_MODE,
            )
        for name in file_names:
            _harden_existing_regular_file(current_path / name)


@contextmanager
def _exclusive_output_directory(out_dir: Path) -> Iterator[None]:
    """Own the stable output-directory inode for freshness check and inference."""

    if fcntl is None:  # pragma: no cover - the submitted runtime is Linux.
        raise V2Error("FormulaBench v2 output locking requires a POSIX runtime")
    if out_dir.is_symlink():
        raise V2Error("output directory must not be a symbolic link")
    try:
        out_dir.mkdir(mode=_PRIVATE_DIRECTORY_MODE, parents=True, exist_ok=True)
        flags = (
            os.O_RDONLY
            | getattr(os, "O_CLOEXEC", 0)
            | getattr(os, "O_DIRECTORY", 0)
            | getattr(os, "O_NOFOLLOW", 0)
        )
        descriptor = os.open(out_dir, flags)
    except OSError as exc:
        raise V2Error("output directory cannot be opened for exclusive ownership") from exc

    locked = False
    try:
        descriptor_stat = os.fstat(descriptor)
        path_stat = os.stat(out_dir, follow_symlinks=False)
        if (
            not stat.S_ISDIR(descriptor_stat.st_mode)
            or not stat.S_ISDIR(path_stat.st_mode)
            or (descriptor_stat.st_dev, descriptor_stat.st_ino)
            != (path_stat.st_dev, path_stat.st_ino)
        ):
            raise V2Error("output directory changed while acquiring exclusive ownership")
        try:
            fcntl.flock(descriptor, fcntl.LOCK_EX | fcntl.LOCK_NB)
        except BlockingIOError as exc:
            raise V2Error("another FormulaBench v2 run already owns the output directory") from exc
        except OSError as exc:
            raise V2Error("output directory lock could not be acquired") from exc
        locked = True
        yield
    finally:
        if locked:
            with suppress(OSError):
                fcntl.flock(descriptor, fcntl.LOCK_UN)
        with suppress(OSError):
            os.close(descriptor)


def _require_fresh_output(out_dir: Path) -> Path:
    # formulabench.artifacts imports openpyxl, so this import must stay behind
    # the v2 serializer configuration in main().
    from .artifacts import ArtifactError, require_writable_directory

    try:
        return require_writable_directory(out_dir)
    except ArtifactError as exc:
        raise V2Error(str(exc)) from exc


def _legacy_main(argv: Sequence[str]) -> int:
    from .capture import main as capture_main

    return capture_main(argv)


def _forward_legacy(argv: list[str]) -> int | None:
    occurrences = argv.count("--legacy-engine")
    if not occurrences:
        return None
    if occurrences > 1:
        parser = argparse.ArgumentParser(prog="python -m formulabench.v2", add_help=False)
        parser.error("--legacy-engine may be specified only once")
    if any(
        argument == "--sampler-checkpoint" or argument.startswith("--sampler-checkpoint=")
        for argument in argv
    ):
        parser = argparse.ArgumentParser(prog="python -m formulabench.v2", add_help=False)
        parser.error(
            "--sampler-checkpoint is not supported by FormulaBench v2 or its "
            "--legacy-engine adapter; use `python -m formulabench.cli` directly "
            "for checkpoint experiments"
        )
    forwarded = list(argv)
    forwarded.remove("--legacy-engine")
    return _legacy_main(forwarded)


def main(argv: Sequence[str] | None = None) -> int:
    raw_argv = list(sys.argv[1:] if argv is None else argv)
    legacy_result = _forward_legacy(raw_argv)
    if legacy_result is not None:
        return legacy_result

    args = parse_args(raw_argv)
    try:
        dataset_dir, out_dir = _validated_paths(args.dataset_dir, args.out_dir)
        _require_canonical_runtime()
        if not args.preflight_only:
            _require_api_key()
        tasks = _load_tasks(dataset_dir)
        if not tasks:
            raise V2Error("dataset contains no tasks")
        selected = _select_tasks(tasks, args.ids)
        if not selected:
            raise V2Error("no tasks selected")

        if args.preflight_only:
            _preflight(tasks, selected)
            return 0

        exactsource_argv = [
            "--data-dir",
            os.fspath(dataset_dir),
            "--out-dir",
            os.fspath(out_dir),
        ]
        if args.ids is not None:
            exactsource_argv.extend(("--ids", args.ids))

        # The directory inode remains stable while freshness is checked and
        # throughout every paid call. ExactSource replaces run.log atomically,
        # so that file cannot serve as the v2 ownership lock.
        with _exclusive_output_directory(out_dir):
            _require_fresh_output(out_dir)
            _harden_output_permissions(out_dir)
            try:
                return _run_exactsource(exactsource_argv)
            finally:
                _harden_output_permissions(out_dir)
    except (V2Error, OSError, ValueError) as exc:
        print(f"FormulaBench v2 startup failed: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
