from __future__ import annotations

import os
import stat
import subprocess
import sys
from pathlib import Path
from types import SimpleNamespace

import pytest

import formulabench.v2 as v2


@pytest.fixture(autouse=True)
def _canonical_runtime(monkeypatch: pytest.MonkeyPatch) -> None:
    """Keep unit tests independent of optional packages in the developer venv."""

    monkeypatch.setattr(v2, "_pillow_available", lambda: False)
    monkeypatch.setattr(v2, "_openpyxl_uses_lxml", lambda: False)
    monkeypatch.setenv("TINKER_API_KEY", "test-only-key")


def _tasks(*task_ids: str) -> list[SimpleNamespace]:
    return [SimpleNamespace(id=task_id) for task_id in task_ids]


def test_v2_dispatches_to_exactsource_with_canonical_paths(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    dataset = tmp_path / "dataset"
    dataset.mkdir()
    out = tmp_path / "fresh-output"
    observed: list[list[str]] = []

    monkeypatch.setattr(v2, "_load_tasks", lambda _: _tasks("first", "second"))

    def fake_run(argv: list[str]) -> int:
        observed.append(argv)
        assert out.is_dir()
        assert not any(out.iterdir())
        return 17

    monkeypatch.setattr(v2, "_run_exactsource", fake_run)

    result = v2.main(
        [
            "--dataset-dir",
            str(dataset),
            "--out-dir",
            str(out),
            "--ids",
            "second,first",
        ]
    )

    assert result == 17
    assert observed == [
        [
            "--data-dir",
            str(dataset.resolve()),
            "--out-dir",
            str(out.resolve()),
            "--ids",
            "second,first",
        ]
    ]


def test_v2_task_selection_matches_exactsource_manifest_order() -> None:
    tasks = _tasks("first", "second", "third")

    assert [task.id for task in v2._select_tasks(tasks, "third,first")] == [
        "first",
        "third",
    ]
    assert [task.id for task in v2._select_tasks(tasks, "first,first")] == ["first"]

    with pytest.raises(v2.V2Error, match="at least one"):
        v2._select_tasks(tasks, ", ,")
    with pytest.raises(v2.V2Error, match="unknown task ids"):
        v2._select_tasks(tasks, "missing")


def test_v2_rejects_overlapping_dataset_and_output(tmp_path: Path) -> None:
    dataset = tmp_path / "dataset"
    dataset.mkdir()

    with pytest.raises(v2.V2Error, match="must not overlap"):
        v2._validated_paths(str(dataset), str(dataset / "out"))
    with pytest.raises(v2.V2Error, match="must not overlap"):
        v2._validated_paths(str(dataset), str(tmp_path))


def test_v2_rejects_output_root_leaf_symlink_before_resolving(tmp_path: Path) -> None:
    dataset = tmp_path / "dataset"
    target = tmp_path / "target"
    supplied = tmp_path / "supplied-output"
    dataset.mkdir()
    target.mkdir()
    supplied.symlink_to(target, target_is_directory=True)

    with pytest.raises(v2.V2Error, match="must not be a symbolic link"):
        v2._validated_paths(str(dataset), str(supplied))

    assert supplied.is_symlink()
    assert not any(target.iterdir())


def test_v2_requires_a_fresh_empty_output_root(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]
) -> None:
    dataset = tmp_path / "dataset"
    dataset.mkdir()
    out = tmp_path / "old-output"
    out.mkdir()
    (out / "predictions.jsonl").write_text("old artefact\n", encoding="utf-8")
    called = False

    monkeypatch.setattr(v2, "_load_tasks", lambda _: _tasks("one"))

    def fake_run(_argv: list[str]) -> int:
        nonlocal called
        called = True
        return 0

    monkeypatch.setattr(v2, "_run_exactsource", fake_run)

    assert v2.main(["--dataset-dir", str(dataset), "--out-dir", str(out)]) == 2
    assert called is False
    assert (out / "predictions.jsonl").read_text(encoding="utf-8") == "old artefact\n"
    assert "output directory is not empty" in capsys.readouterr().err


def test_v2_preflight_loads_and_selects_without_provider_or_output_calls(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]
) -> None:
    dataset = tmp_path / "dataset"
    dataset.mkdir()
    out = tmp_path / "must-not-be-created"
    loaded: list[Path] = []
    monkeypatch.delenv("TINKER_API_KEY")

    def fake_load(path: Path) -> list[SimpleNamespace]:
        loaded.append(path)
        return _tasks("one", "two")

    monkeypatch.setattr(v2, "_load_tasks", fake_load)
    monkeypatch.setattr(
        v2,
        "_run_exactsource",
        lambda _argv: pytest.fail("preflight made an ExactSource/provider call"),
    )
    monkeypatch.setattr(
        v2,
        "_require_fresh_output",
        lambda _path: pytest.fail("preflight touched the output root"),
    )

    assert (
        v2.main(
            [
                "--dataset-dir",
                str(dataset),
                "--out-dir",
                str(out),
                "--ids",
                "two",
                "--preflight-only",
            ]
        )
        == 0
    )

    assert loaded == [dataset.resolve()]
    assert not out.exists()
    output = capsys.readouterr().out
    assert "tasks=2" in output
    assert "selected=1" in output
    assert "engine=ExactSource/0.1.0" in output
    assert "model=tinker:Qwen/Qwen3.8-27B" in output
    assert "vendored_source=ExactSource@99fe8084bf35a5fca6a2c2e1c9beae802766a618" in output
    assert "scored_core=ExactSource@8b84dba1d9263e2123b8f15267239b70ff817907" in output


@pytest.mark.parametrize("credential", [None, "", " \t\n"])
def test_v2_rejects_missing_or_blank_key_before_loading_tasks_or_creating_output(
    credential: str | None,
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    capsys: pytest.CaptureFixture[str],
) -> None:
    dataset = tmp_path / "dataset"
    out = tmp_path / "must-not-be-created"
    dataset.mkdir()
    if credential is None:
        monkeypatch.delenv("TINKER_API_KEY", raising=False)
    else:
        monkeypatch.setenv("TINKER_API_KEY", credential)
    monkeypatch.setattr(
        v2,
        "_load_tasks",
        lambda _path: pytest.fail("missing credential loaded tasks"),
    )
    monkeypatch.setattr(
        v2,
        "_run_exactsource",
        lambda _argv: pytest.fail("missing credential called the provider"),
    )

    assert v2.main(["--dataset-dir", str(dataset), "--out-dir", str(out)]) == 2

    assert not out.exists()
    error = capsys.readouterr().err
    assert "TINKER_API_KEY is not set" in error
    assert "test-only-key" not in error


@pytest.mark.parametrize(
    ("guard", "message"),
    [
        ("_pillow_available", "Pillow/PIL is installed"),
        ("_openpyxl_uses_lxml", "openpyxl is using its lxml serializer"),
    ],
)
@pytest.mark.parametrize(
    "extra_arguments",
    [[], ["--preflight-only"]],
    ids=["paid-run", "preflight"],
)
def test_v2_fails_closed_on_noncanonical_dependencies_before_provider_or_output(
    guard: str,
    message: str,
    extra_arguments: list[str],
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    capsys: pytest.CaptureFixture[str],
) -> None:
    dataset = tmp_path / "dataset"
    out = tmp_path / "must-not-be-created"
    dataset.mkdir()
    monkeypatch.setattr(v2, guard, lambda: True)
    monkeypatch.setattr(
        v2,
        "_load_tasks",
        lambda _path: pytest.fail("dependency failure loaded tasks"),
    )
    monkeypatch.setattr(
        v2,
        "_run_exactsource",
        lambda _argv: pytest.fail("dependency failure called the provider"),
    )

    assert v2.main(["--dataset-dir", str(dataset), "--out-dir", str(out), *extra_arguments]) == 2

    assert not out.exists()
    error = capsys.readouterr().err
    assert message in error
    assert "Docker runner" in error
    assert "uv sync --locked" in error


@pytest.mark.skipif(v2.fcntl is None, reason="v2 output locking requires POSIX")
def test_v2_rejects_concurrent_owner_before_freshness_or_provider(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    capsys: pytest.CaptureFixture[str],
) -> None:
    dataset = tmp_path / "dataset"
    out = tmp_path / "output"
    dataset.mkdir()
    monkeypatch.setattr(v2, "_load_tasks", lambda _: _tasks("one"))
    monkeypatch.setattr(
        v2,
        "_require_fresh_output",
        lambda _path: pytest.fail("contending run checked freshness"),
    )
    monkeypatch.setattr(
        v2,
        "_run_exactsource",
        lambda _argv: pytest.fail("contending run called the provider"),
    )

    with v2._exclusive_output_directory(out):
        result = v2.main(["--dataset-dir", str(dataset), "--out-dir", str(out)])

    assert result == 2
    assert "already owns the output directory" in capsys.readouterr().err


@pytest.mark.skipif(v2.fcntl is None, reason="v2 output locking requires POSIX")
def test_v2_holds_output_inode_lock_for_entire_provider_run(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    dataset = tmp_path / "dataset"
    out = tmp_path / "output"
    dataset.mkdir()
    events: list[str] = []
    monkeypatch.setattr(v2, "_load_tasks", lambda _: _tasks("one"))

    def assert_owned(phase: str) -> None:
        with (
            pytest.raises(v2.V2Error, match="already owns the output directory"),
            v2._exclusive_output_directory(out),
        ):
            pass
        events.append(phase)

    def fake_fresh(path: Path) -> Path:
        assert_owned("freshness")
        return path

    def fake_run(_argv: list[str]) -> int:
        assert_owned("provider")
        return 0

    monkeypatch.setattr(v2, "_require_fresh_output", fake_fresh)
    monkeypatch.setattr(v2, "_run_exactsource", fake_run)

    assert v2.main(["--dataset-dir", str(dataset), "--out-dir", str(out)]) == 0
    assert events == ["freshness", "provider"]


def test_v2_hardens_successful_output_tree(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    dataset = tmp_path / "dataset"
    out = tmp_path / "output"
    dataset.mkdir()
    monkeypatch.setattr(v2, "_load_tasks", lambda _: _tasks("one"))

    def fake_run(_argv: list[str]) -> int:
        assert stat.S_IMODE(out.stat().st_mode) == 0o700
        outputs = out / "outputs"
        outputs.mkdir(mode=0o755)
        workbook = outputs / "one.xlsx"
        workbook.write_bytes(b"workbook")
        workbook.chmod(0o644)
        return 0

    monkeypatch.setattr(v2, "_run_exactsource", fake_run)

    assert v2.main(["--dataset-dir", str(dataset), "--out-dir", str(out)]) == 0
    assert stat.S_IMODE(out.stat().st_mode) == 0o700
    assert stat.S_IMODE((out / "outputs").stat().st_mode) == 0o700
    assert stat.S_IMODE((out / "outputs" / "one.xlsx").stat().st_mode) == 0o600


def test_v2_hardens_partial_output_when_run_fails(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    capsys: pytest.CaptureFixture[str],
) -> None:
    dataset = tmp_path / "dataset"
    out = tmp_path / "output"
    dataset.mkdir()
    monkeypatch.setattr(v2, "_load_tasks", lambda _: _tasks("one"))

    def fake_run(_argv: list[str]) -> int:
        traces = out / "traces"
        traces.mkdir(mode=0o755)
        trace = traces / "one.jsonl"
        trace.write_text("partial\n", encoding="utf-8")
        trace.chmod(0o644)
        raise OSError("simulated run failure")

    monkeypatch.setattr(v2, "_run_exactsource", fake_run)

    assert v2.main(["--dataset-dir", str(dataset), "--out-dir", str(out)]) == 2
    assert "simulated run failure" in capsys.readouterr().err
    assert stat.S_IMODE(out.stat().st_mode) == 0o700
    assert stat.S_IMODE((out / "traces").stat().st_mode) == 0o700
    assert stat.S_IMODE((out / "traces" / "one.jsonl").stat().st_mode) == 0o600


@pytest.mark.parametrize("value", ["0", "1", "3", "5", "64", "not-an-int"])
def test_v2_accepts_only_concurrency_four(value: str) -> None:
    with pytest.raises(SystemExit):
        v2.parse_args(["--dataset-dir", "/data", "--out-dir", "/out", "--concurrency", value])

    assert (
        v2.parse_args(
            ["--dataset-dir", "/data", "--out-dir", "/out", "--concurrency", "4"]
        ).concurrency
        == 4
    )


def test_v2_help_exposes_legacy_engine_and_fixed_concurrency(
    capsys: pytest.CaptureFixture[str],
) -> None:
    with pytest.raises(SystemExit) as raised:
        v2.parse_args(["--help"])

    assert raised.value.code == 0
    help_text = " ".join(capsys.readouterr().out.split())
    assert "--legacy-engine" in help_text
    assert "run the original FormulaBench engine instead of v2" in help_text
    assert "--concurrency CONCURRENCY" in help_text
    assert "fixed at 4 for ExactSource parity" in help_text


@pytest.mark.parametrize("abbreviation", ["--preflight", "--preflight-o", "--legacy-e"])
def test_v2_rejects_abbreviated_mode_flags(
    abbreviation: str, capsys: pytest.CaptureFixture[str]
) -> None:
    with pytest.raises(SystemExit) as raised:
        v2.parse_args(["--dataset-dir", "/data", "--out-dir", "/out", abbreviation])

    assert raised.value.code == 2
    assert f"unrecognized arguments: {abbreviation}" in capsys.readouterr().err


@pytest.mark.parametrize(
    ("option", "arguments"),
    [
        ("--resume", ["--resume"]),
        ("--retry-failures", ["--retry-failures"]),
        ("--replay-write-failures", ["--replay-write-failures"]),
    ],
)
def test_v2_rejects_resume_style_options_with_legacy_direction(
    option: str, arguments: list[str], capsys: pytest.CaptureFixture[str]
) -> None:
    with pytest.raises(SystemExit):
        v2.parse_args(["--dataset-dir", "/data", "--out-dir", "/out", *arguments])

    error = capsys.readouterr().err
    assert f"{option} is not supported by FormulaBench v2" in error
    assert "use --legacy-engine" in error


@pytest.mark.parametrize("legacy", [False, True], ids=["v2", "legacy-adapter"])
def test_v2_rejects_sampler_checkpoint_with_direct_cli_direction(
    legacy: bool,
    monkeypatch: pytest.MonkeyPatch,
    capsys: pytest.CaptureFixture[str],
) -> None:
    monkeypatch.setattr(
        v2,
        "_legacy_main",
        lambda _argv: pytest.fail("unsupported checkpoint reached the legacy adapter"),
    )
    arguments = ["--dataset-dir", "/data", "--out-dir", "/out"]
    if legacy:
        arguments.append("--legacy-engine")
    arguments.extend(("--sampler-checkpoint", "tinker://checkpoint"))

    with pytest.raises(SystemExit):
        v2.main(arguments)

    error = capsys.readouterr().err
    assert "--sampler-checkpoint is not supported" in error
    assert "FormulaBench v2 or its --legacy-engine adapter" in error
    assert "python -m formulabench.cli" in error
    assert "use --legacy-engine to run the legacy engine" not in error


def test_legacy_engine_forwards_without_changing_serializer_environment(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    forwarded: list[list[str]] = []
    monkeypatch.setenv("OPENPYXL_LXML", "True")

    def fake_legacy(argv: list[str]) -> int:
        forwarded.append(argv)
        return 23

    monkeypatch.setattr(v2, "_legacy_main", fake_legacy)
    original = [
        "--dataset-dir",
        "/data",
        "--legacy-engine",
        "--out-dir",
        "/old-out",
        "--resume",
        "--retry-failures",
    ]

    assert v2.main(original) == 23
    assert os.environ["OPENPYXL_LXML"] == "True"
    assert forwarded == [
        [
            "--dataset-dir",
            "/data",
            "--out-dir",
            "/old-out",
            "--resume",
            "--retry-failures",
        ]
    ]


def test_clean_v2_process_disables_openpyxl_lxml_before_import() -> None:
    environment = dict(os.environ)
    environment.pop("OPENPYXL_LXML", None)
    script = """
import sys
import formulabench.v2 as v2
assert "openpyxl" not in sys.modules
assert "exactsource" not in sys.modules
try:
    v2.parse_args(["--version"])
except SystemExit as error:
    assert error.code == 0
import openpyxl.xml.functions
print(f"OPENPYXL_LXML={openpyxl.xml.functions.LXML}")
"""

    completed = subprocess.run(
        [sys.executable, "-c", script],
        cwd=Path(__file__).resolve().parents[1],
        env=environment,
        check=True,
        capture_output=True,
        text=True,
    )

    assert completed.stdout.splitlines()[-1] == "OPENPYXL_LXML=False"


def test_exactsource_dispatch_installs_runtime_adapters_before_run(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import exactsource.cli

    events: list[str] = []
    monkeypatch.setattr(v2, "_configure_v2_serializer", lambda: events.append("configure-parent"))
    monkeypatch.setattr(v2, "_require_canonical_runtime", lambda: events.append("check-runtime"))
    monkeypatch.setattr(
        v2,
        "_install_exactsource_private_permissions",
        lambda: events.append("install-permissions"),
    )
    monkeypatch.setattr(
        v2,
        "_install_exactsource_child_serializer",
        lambda: events.append("install-child-hook"),
    )
    monkeypatch.setattr(
        exactsource.cli,
        "run_cli",
        lambda _argv: events.append("run-exactsource") or 31,
    )

    assert v2._run_exactsource(["--data-dir", "/data", "--out-dir", "/out"]) == 31
    assert events == [
        "configure-parent",
        "check-runtime",
        "install-permissions",
        "install-child-hook",
        "run-exactsource",
    ]


def test_exactsource_permission_adapter_makes_published_artifacts_private(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    import openpyxl

    from exactsource import artifacts, plans, sandbox

    # Register the original state with monkeypatch so this v2-only adapter does
    # not leak into unrelated ExactSource tests in the same pytest process.
    monkeypatch.setattr(artifacts, "_PUBLISHED_DIRECTORY_MODE", artifacts._PUBLISHED_DIRECTORY_MODE)
    monkeypatch.setattr(artifacts, "_PUBLISHED_FILE_MODE", artifacts._PUBLISHED_FILE_MODE)
    monkeypatch.setattr(plans, "_save_atomic", plans._save_atomic)
    monkeypatch.setattr(sandbox, "_promote_atomic", sandbox._promote_atomic)

    v2._install_exactsource_private_permissions()
    layout = artifacts.prepare_output(tmp_path / "output")
    artifacts.atomic_write_text(layout.log_path, "private\n")

    workbook = openpyxl.Workbook()
    operation_output = layout.outputs_dir / "operation.xlsx"
    plans._save_atomic(workbook, operation_output)
    workbook.close()

    source = tmp_path / "source.xlsx"
    source.write_bytes(operation_output.read_bytes())
    python_output = layout.outputs_dir / "python.xlsx"
    sandbox._promote_atomic(source, python_output)

    for directory in (layout.root, layout.outputs_dir, layout.traces_dir):
        assert stat.S_IMODE(directory.stat().st_mode) == 0o700
    for file_path in (layout.log_path, operation_output, python_output):
        assert stat.S_IMODE(file_path.stat().st_mode) == 0o600


def test_clean_v2_python_transform_uses_stdlib_xml_in_child(tmp_path: Path) -> None:
    environment = dict(os.environ)
    environment.pop("OPENPYXL_LXML", None)
    script = """
import sys
import zipfile
from pathlib import Path

import formulabench.v2 as v2
assert "openpyxl" not in sys.modules
v2._configure_v2_serializer()
v2._install_exactsource_child_serializer()

import openpyxl
from exactsource import sandbox

installed_builder = sandbox._child_environment
v2._install_exactsource_child_serializer()
assert sandbox._child_environment is installed_builder

root = Path(sys.argv[1])
source = root / "source.xlsx"
destination = root / "transformed.xlsx"
workbook = openpyxl.Workbook()
workbook.active["A1"] = "before"
workbook.save(source)
workbook.close()

sandbox.run_transform(
    "def transform(wb):\\n    wb.active['A1'] = 'after'\\n",
    source,
    destination,
)
with zipfile.ZipFile(destination) as archive:
    content_types = archive.read("[Content_Types].xml")

assert sandbox._child_environment(1)["OPENPYXL_LXML"] == "False"
relationship_type = b'ContentType="application/vnd.openxmlformats-package.relationships+xml"'
assert relationship_type + b" />" in content_types
assert relationship_type + b"/>" not in content_types
print("child_serializer=stdlib")
"""

    completed = subprocess.run(
        [sys.executable, "-c", script, str(tmp_path)],
        cwd=Path(__file__).resolve().parents[1],
        env=environment,
        check=True,
        capture_output=True,
        text=True,
    )

    assert completed.stdout.strip() == "child_serializer=stdlib"


def test_v2_version_identifies_exactsource(capsys: pytest.CaptureFixture[str]) -> None:
    with pytest.raises(SystemExit) as raised:
        v2.parse_args(["--version"])

    assert raised.value.code == 0
    assert " ".join(capsys.readouterr().out.splitlines()) == (
        "python -m formulabench.v2 ExactSource/0.1.0 "
        "vendored_source=ExactSource@99fe8084bf35a5fca6a2c2e1c9beae802766a618 "
        "scored_core=ExactSource@8b84dba1d9263e2123b8f15267239b70ff817907"
    )
