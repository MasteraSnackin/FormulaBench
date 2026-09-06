from __future__ import annotations

import os
import subprocess
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
EVALUATE_CHECKPOINT = ROOT / "scripts" / "evaluate_checkpoint.sh"
RUN_TRAINING = ROOT / "scripts" / "run_training.sh"


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
