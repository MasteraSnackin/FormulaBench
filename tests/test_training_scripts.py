from __future__ import annotations

import os
import subprocess
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
EVALUATE_CHECKPOINT = ROOT / "scripts" / "evaluate_checkpoint.sh"
RUN_TRAINING = ROOT / "scripts" / "run_training.sh"


def _recording_uv(tmp_path: Path) -> tuple[Path, Path]:
    bin_dir = tmp_path / "bin"
    bin_dir.mkdir()
    log = tmp_path / "uv-environment.log"
    executable = bin_dir / "uv"
    executable.write_text(
        "#!/bin/sh\n"
        'if [ -n "${TINKER_API_KEY:-}" ]; then key_state=set; else key_state=unset; fi\n'
        'printf \'%s|%s\\n\' "$key_state" "$*" >> "$UV_ENVIRONMENT_LOG"\n'
        "case \"$*\" in *'training.partition_ids'*) printf 'task-1\\n' ;; esac\n",
        encoding="utf-8",
    )
    executable.chmod(0o755)
    return bin_dir, log


@pytest.mark.parametrize(
    "override",
    [
        ["--manifest", "/tmp/alternate.json"],
        ["--manifest=/tmp/alternate.json"],
        ["--split-manifest", "/tmp/alternate-split.json"],
        ["--split-manifest=/tmp/alternate-split.json"],
    ],
)
def test_training_wrapper_refuses_manifest_overrides(override: list[str]) -> None:
    result = subprocess.run(
        [str(RUN_TRAINING), *override],
        cwd=ROOT,
        capture_output=True,
        check=False,
        text=True,
    )

    assert result.returncode == 2
    assert result.stdout == ""
    assert result.stderr == (
        "run_training.sh fixes the corpus and split manifests; overrides are not allowed\n"
    )


def test_docker_context_excludes_local_training_evidence() -> None:
    patterns = {
        line.strip() for line in (ROOT / ".dockerignore").read_text(encoding="utf-8").splitlines()
    }

    assert {"training/generated/", "training/runs/"}.issubset(patterns)


def test_checkpoint_evaluation_requires_explicit_paid_authorisation(tmp_path: Path) -> None:
    result = subprocess.run(
        [str(EVALUATE_CHECKPOINT)],
        cwd=ROOT,
        capture_output=True,
        check=False,
        text=True,
    )

    assert result.returncode == 2
    assert "usage:" in result.stderr
    assert not any(tmp_path.iterdir())


def test_checkpoint_evaluation_requires_key_before_creating_output(tmp_path: Path) -> None:
    output_root = tmp_path / "evaluation"
    environment = dict(os.environ)
    environment.pop("TINKER_API_KEY", None)

    result = subprocess.run(
        [str(EVALUATE_CHECKPOINT), "--execute", "tinker://test", str(output_root)],
        cwd=ROOT,
        env=environment,
        capture_output=True,
        check=False,
        text=True,
    )

    assert result.returncode == 2
    assert result.stderr == (
        "TINKER_API_KEY is required and must be supplied through the environment\n"
    )
    assert not output_root.exists()


def test_checkpoint_evaluation_requires_explicit_project_before_output(tmp_path: Path) -> None:
    output_root = tmp_path / "evaluation"
    environment = {**os.environ, "TINKER_API_KEY": "test-only-key"}
    environment.pop("TINKER_PROJECT_ID", None)

    result = subprocess.run(
        [
            str(EVALUATE_CHECKPOINT),
            "--execute",
            "tinker://test/sampler_weights/checkpoint",
            str(output_root),
        ],
        cwd=ROOT,
        env=environment,
        capture_output=True,
        check=False,
        text=True,
    )

    assert result.returncode == 2
    assert result.stderr == (
        "TINKER_PROJECT_ID is required and must identify the checkpoint's writable project\n"
    )
    assert not output_root.exists()


def test_training_wrapper_limits_key_to_paid_trainer(tmp_path: Path) -> None:
    bin_dir, log = _recording_uv(tmp_path)
    environment = {
        **os.environ,
        "PATH": f"{bin_dir}:{os.environ['PATH']}",
        "TINKER_API_KEY": "test-only-key",
        "TINKER_PROJECT_ID": "test-project",
        "UV_ENVIRONMENT_LOG": str(log),
    }

    result = subprocess.run(
        [
            str(RUN_TRAINING),
            "--execute",
            "--project-id",
            "test-project",
            "--run-dir",
            str(tmp_path / "run"),
        ],
        cwd=ROOT,
        env=environment,
        capture_output=True,
        check=False,
        text=True,
    )

    assert result.returncode == 0
    records = log.read_text(encoding="utf-8").splitlines()
    assert len(records) == 4
    assert all(record.startswith("unset|") for record in records[:3])
    assert records[3].startswith("set|run --extra native-tinker python -m training.train_lora ")


def test_checkpoint_wrapper_limits_key_to_paid_inference(tmp_path: Path) -> None:
    bin_dir, log = _recording_uv(tmp_path)
    output_root = tmp_path / "evaluation"
    environment = {
        **os.environ,
        "PATH": f"{bin_dir}:{os.environ['PATH']}",
        "TINKER_API_KEY": "test-only-key",
        "TINKER_PROJECT_ID": "test-project",
        "UV_ENVIRONMENT_LOG": str(log),
    }

    result = subprocess.run(
        [
            str(EVALUATE_CHECKPOINT),
            "--execute",
            "tinker://test/sampler_weights/checkpoint",
            str(output_root),
            str(tmp_path / "dataset"),
        ],
        cwd=ROOT,
        env=environment,
        capture_output=True,
        check=False,
        text=True,
    )

    assert result.returncode == 0
    records = log.read_text(encoding="utf-8").splitlines()
    paid = [record for record in records if record.startswith("set|")]
    local = [record for record in records if record.startswith("unset|")]
    assert len(paid) == 2
    assert all("python -m formulabench.cli" in record for record in paid)
    assert len(local) == len(records) - 2
    assert any("python -m training.corpus build" in record for record in local)
    assert any("python evaluate.py" in record for record in local)
