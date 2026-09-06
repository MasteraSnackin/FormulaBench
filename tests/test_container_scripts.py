from __future__ import annotations

import os
import subprocess
from pathlib import Path

import pytest

PROJECT_ROOT = Path(__file__).resolve().parents[1]
RUN_DOCKER = PROJECT_ROOT / "scripts" / "run_docker.sh"
CONTAINER_ENTRYPOINT = PROJECT_ROOT / "scripts" / "container_entrypoint.sh"


def _fake_docker(tmp_path: Path) -> Path:
    executable = tmp_path / "bin" / "docker"
    executable.parent.mkdir()
    executable.write_text(
        "#!/bin/sh\n"
        "printf 'DOCKER_CALL\\n'\n"
        'for argument in "$@"; do\n'
        "    printf 'DOCKER_ARG=%s\\n' \"$argument\"\n"
        "done\n",
        encoding="utf-8",
    )
    executable.chmod(0o755)
    return executable


def _wrapper_fixture(tmp_path: Path) -> tuple[Path, Path, dict[str, str]]:
    dataset = tmp_path / "dataset"
    dataset.mkdir()
    (dataset / "dataset.json").write_text("[]\n", encoding="utf-8")
    output = tmp_path / "output"
    fake_docker = _fake_docker(tmp_path)
    environment = os.environ.copy()
    environment["PATH"] = f"{fake_docker.parent}{os.pathsep}{environment['PATH']}"
    environment["TINKER_API_KEY"] = "must-not-appear-in-docker-arguments"
    environment["TINKER_PROJECT_ID"] = "must-not-appear-in-docker-arguments"
    return dataset, output, environment


@pytest.mark.parametrize(
    "override",
    [
        "--dataset-dir",
        "--dataset-dir=/different-data",
        "--out-dir",
        "--out-dir=/different-output",
    ],
)
def test_docker_wrapper_rejects_container_mount_overrides(override: str, tmp_path: Path) -> None:
    completed = subprocess.run(
        [str(RUN_DOCKER), str(tmp_path / "dataset"), str(tmp_path / "output"), override],
        cwd=PROJECT_ROOT,
        capture_output=True,
        text=True,
        check=False,
    )

    assert completed.returncode == 2
    assert "--dataset-dir and --out-dir are managed by this wrapper" in completed.stderr
    assert "first two positional arguments" in completed.stderr


@pytest.mark.parametrize(
    "arguments",
    [
        ["--dataset-dir=/different-data", "output"],
        ["dataset", "--out-dir=/different-output"],
    ],
)
def test_docker_wrapper_rejects_mount_flags_in_positional_slots(arguments: list[str]) -> None:
    completed = subprocess.run(
        [str(RUN_DOCKER), *arguments],
        cwd=PROJECT_ROOT,
        capture_output=True,
        text=True,
        check=False,
    )

    assert completed.returncode == 2
    assert "--dataset-dir and --out-dir are managed by this wrapper" in completed.stderr
    assert "first two positional arguments" in completed.stderr


def test_docker_wrapper_does_not_forward_credentials_during_preflight(tmp_path: Path) -> None:
    dataset, output, environment = _wrapper_fixture(tmp_path)

    completed = subprocess.run(
        [str(RUN_DOCKER), str(dataset), str(output), "--preflight-only"],
        cwd=PROJECT_ROOT,
        env=environment,
        capture_output=True,
        text=True,
        check=False,
    )

    assert completed.returncode == 0
    assert completed.stdout.count("DOCKER_CALL\n") == 2
    assert "DOCKER_ARG=TINKER_API_KEY\n" not in completed.stdout
    assert "DOCKER_ARG=TINKER_PROJECT_ID\n" not in completed.stdout


def test_docker_wrapper_rejects_whitespace_only_api_key_before_output_creation(
    tmp_path: Path,
) -> None:
    dataset, output, environment = _wrapper_fixture(tmp_path)
    environment["TINKER_API_KEY"] = " \t\n"

    completed = subprocess.run(
        [str(RUN_DOCKER), str(dataset), str(output)],
        cwd=PROJECT_ROOT,
        env=environment,
        capture_output=True,
        text=True,
        check=False,
    )

    assert completed.returncode == 2
    assert "TINKER_API_KEY is not set" in completed.stderr
    assert not output.exists()
    assert "DOCKER_CALL" not in completed.stdout


def test_docker_wrapper_rejects_new_output_nested_inside_dataset(tmp_path: Path) -> None:
    dataset, _output, environment = _wrapper_fixture(tmp_path)
    output = dataset / "generated-output"

    completed = subprocess.run(
        [str(RUN_DOCKER), str(dataset), str(output), "--preflight-only"],
        cwd=PROJECT_ROOT,
        env=environment,
        capture_output=True,
        text=True,
        check=False,
    )

    assert completed.returncode == 2
    assert "Dataset and output directories must not overlap" in completed.stderr
    assert "DOCKER_CALL" not in completed.stdout
    assert not output.exists()
    assert {path.name for path in dataset.iterdir()} == {"dataset.json"}


def test_docker_wrapper_rejects_multilevel_output_without_mutating_dataset(
    tmp_path: Path,
) -> None:
    dataset, _output, environment = _wrapper_fixture(tmp_path)
    output_parent = dataset / "must-not-be-created"
    output = output_parent / "nested-output"

    completed = subprocess.run(
        [str(RUN_DOCKER), str(dataset), str(output), "--preflight-only"],
        cwd=PROJECT_ROOT,
        env=environment,
        capture_output=True,
        text=True,
        check=False,
    )

    assert completed.returncode == 2
    assert "Dataset and output directories must not overlap" in completed.stderr
    assert "DOCKER_CALL" not in completed.stdout
    assert not output_parent.exists()
    assert {path.name for path in dataset.iterdir()} == {"dataset.json"}


def test_docker_wrapper_preserves_existing_nested_output_when_rejecting_overlap(
    tmp_path: Path,
) -> None:
    dataset, _output, environment = _wrapper_fixture(tmp_path)
    output = dataset / "existing-output"
    output.mkdir()

    completed = subprocess.run(
        [str(RUN_DOCKER), str(dataset), str(output), "--preflight-only"],
        cwd=PROJECT_ROOT,
        env=environment,
        capture_output=True,
        text=True,
        check=False,
    )

    assert completed.returncode == 2
    assert "Dataset and output directories must not overlap" in completed.stderr
    assert "DOCKER_CALL" not in completed.stdout
    assert output.is_dir()
    assert not any(output.iterdir())


@pytest.mark.parametrize("suffix", ["/", "/."])
def test_docker_wrapper_rejects_semantic_output_leaf_symlink(
    suffix: str,
    tmp_path: Path,
) -> None:
    dataset, _output, environment = _wrapper_fixture(tmp_path)
    target = tmp_path / "symlink-target"
    target.mkdir()
    supplied = tmp_path / "output-link"
    supplied.symlink_to(target, target_is_directory=True)

    completed = subprocess.run(
        [str(RUN_DOCKER), str(dataset), f"{supplied}{suffix}", "--preflight-only"],
        cwd=PROJECT_ROOT,
        env=environment,
        capture_output=True,
        text=True,
        check=False,
    )

    assert completed.returncode == 2
    assert "Output directory must not be a symbolic link" in completed.stderr
    assert "DOCKER_CALL" not in completed.stdout
    assert supplied.is_symlink()
    assert target.is_dir()
    assert not any(target.iterdir())


def test_docker_wrapper_accepts_valid_non_overlapping_output(tmp_path: Path) -> None:
    dataset, output, environment = _wrapper_fixture(tmp_path)

    completed = subprocess.run(
        [str(RUN_DOCKER), str(dataset), str(output), "--preflight-only"],
        cwd=PROJECT_ROOT,
        env=environment,
        capture_output=True,
        text=True,
        check=False,
    )

    assert completed.returncode == 0
    assert completed.stdout.count("DOCKER_CALL\n") == 2
    assert f"DOCKER_ARG=type=bind,src={dataset.resolve()},dst=/data,readonly\n" in completed.stdout
    assert f"DOCKER_ARG=type=bind,src={output.resolve()},dst=/out\n" in completed.stdout
    assert output.is_dir()
    assert not any(output.iterdir())


def test_docker_wrapper_forwards_credentials_for_a_paid_run(tmp_path: Path) -> None:
    dataset, output, environment = _wrapper_fixture(tmp_path)

    completed = subprocess.run(
        [str(RUN_DOCKER), str(dataset), str(output)],
        cwd=PROJECT_ROOT,
        env=environment,
        capture_output=True,
        text=True,
        check=False,
    )

    assert completed.returncode == 0
    assert completed.stdout.count("DOCKER_CALL\n") == 2
    assert "DOCKER_ARG=TINKER_API_KEY\n" in completed.stdout
    assert "DOCKER_ARG=TINKER_PROJECT_ID\n" not in completed.stdout


def test_docker_wrapper_forwards_project_id_only_for_paid_legacy_run(tmp_path: Path) -> None:
    dataset, output, environment = _wrapper_fixture(tmp_path)

    completed = subprocess.run(
        [str(RUN_DOCKER), str(dataset), str(output), "--legacy-engine"],
        cwd=PROJECT_ROOT,
        env=environment,
        capture_output=True,
        text=True,
        check=False,
    )

    assert completed.returncode == 0
    assert completed.stdout.count("DOCKER_CALL\n") == 2
    assert "DOCKER_ARG=TINKER_API_KEY\n" in completed.stdout
    assert "DOCKER_ARG=TINKER_PROJECT_ID\n" in completed.stdout


def test_container_entrypoint_puts_fixed_mounts_after_caller_arguments(tmp_path: Path) -> None:
    captured = tmp_path / "arguments.txt"
    recorder = tmp_path / "record-arguments.sh"
    recorder.write_text(
        '#!/bin/sh\nfor argument in "$@"; do\n    printf \'%s\\n\' "$argument"\ndone\n',
        encoding="utf-8",
    )
    recorder.chmod(0o755)

    script = CONTAINER_ENTRYPOINT.read_text(encoding="utf-8").replace(
        "/app/.venv/bin/python", str(recorder), 1
    )
    executable = tmp_path / "container-entrypoint.sh"
    executable.write_text(script, encoding="utf-8")
    executable.chmod(0o755)

    with captured.open("w", encoding="utf-8") as output:
        completed = subprocess.run(
            [
                str(executable),
                "--ids=task-1",
                "--dataset-dir=/caller-data",
                "--out-dir=/caller-output",
            ],
            cwd=PROJECT_ROOT,
            stdout=output,
            text=True,
            check=False,
        )

    assert completed.returncode == 0
    assert captured.read_text(encoding="utf-8").splitlines() == [
        "-m",
        "formulabench.v2",
        "--ids=task-1",
        "--dataset-dir=/caller-data",
        "--out-dir=/caller-output",
        "--dataset-dir=/data",
        "--out-dir=/out",
    ]
