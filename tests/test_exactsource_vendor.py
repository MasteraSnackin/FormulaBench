from __future__ import annotations

import hashlib
import json
from pathlib import Path

from formulabench.v2 import SCORED_CORE_ID, VENDORED_SOURCE_ID

PROJECT_ROOT = Path(__file__).resolve().parents[1]
VENDOR_ROOT = PROJECT_ROOT / "exactsource"
EXPECTED_SCORED_CORE_ID = "ExactSource@8b84dba1d9263e2123b8f15267239b70ff817907"
EXPECTED_VENDORED_SOURCE_ID = "ExactSource@99fe8084bf35a5fca6a2c2e1c9beae802766a618"
EXPECTED_TREE_SHA256 = "293c0958fcdbc86addff96fbc651a8463eb5577063a4a0d3cb5b406b1d16cf00"
EXPECTED_FILES = (
    "__init__.py",
    "artifacts.py",
    "cli.py",
    "config.py",
    "context.py",
    "contracts.py",
    "dataset.py",
    "formula_safety.py",
    "metrics.py",
    "model.py",
    "plans.py",
    "prompts.py",
    "ranges.py",
    "runner.py",
    "sandbox.py",
    "workbook.py",
)


def _vendored_tree_digest() -> str:
    digest = hashlib.sha256()
    for name in EXPECTED_FILES:
        content = (VENDOR_ROOT / name).read_bytes()
        encoded_name = name.encode("utf-8")
        digest.update(len(encoded_name).to_bytes(8, "big"))
        digest.update(encoded_name)
        digest.update(len(content).to_bytes(8, "big"))
        digest.update(content)
    return digest.hexdigest()


def test_vendored_exactsource_tree_matches_recorded_scored_core() -> None:
    actual_files = tuple(path.name for path in sorted(VENDOR_ROOT.glob("*.py")))

    assert actual_files == EXPECTED_FILES
    assert _vendored_tree_digest() == EXPECTED_TREE_SHA256
    assert SCORED_CORE_ID == EXPECTED_SCORED_CORE_ID
    assert VENDORED_SOURCE_ID == EXPECTED_VENDORED_SOURCE_ID


def test_migration_evidence_identifies_the_same_vendored_tree() -> None:
    evidence = json.loads(
        (PROJECT_ROOT / "experiments" / "v2_migration_validation.json").read_text(encoding="utf-8")
    )

    assert evidence["source"]["exactsource_scored_commit"] == EXPECTED_SCORED_CORE_ID.removeprefix(
        "ExactSource@"
    )
    assert evidence["source"]["exactsource_vendored_from_commit"] == (
        EXPECTED_VENDORED_SOURCE_ID.removeprefix("ExactSource@")
    )
    assert evidence["source"]["vendored_python_files"] == len(EXPECTED_FILES)
    assert evidence["source"]["vendored_tree_sha256"] == EXPECTED_TREE_SHA256
    verifier_digest = hashlib.sha256(
        (PROJECT_ROOT / "tools" / "verify_v2_parity.py").read_bytes()
    ).hexdigest()
    assert evidence["credential_free_parity"]["verifier_sha256"] == verifier_digest
    assert evidence["upstream_scored_reference"]["solver_commit"] == (
        EXPECTED_SCORED_CORE_ID.removeprefix("ExactSource@")
    )
    assert evidence["upstream_scored_reference"]["result"] == {
        "tasks": 400,
        "passed": 302,
        "pass_rate": 0.755,
        "cell_accuracy": 0.8006,
        "pass_rate_cell_level": 0.7818,
        "pass_rate_sheet_level": 0.696,
    }
