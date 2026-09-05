from __future__ import annotations

import argparse
import asyncio
import json
from pathlib import Path
from types import SimpleNamespace

import openpyxl
import pytest

import formulabench.capture as capture
import formulabench.cli as cli
from formulabench.artifacts import SUPERVISOR_LOCK_FD_ENV, ArtifactError, DatasetTask
from formulabench.capture import CaptureResult
from formulabench.cli import _ensure_output_outside_dataset, _select_tasks, parse_args
from formulabench.constants import MODEL_ID
from formulabench.provider import Completion


def _task(task_id: str) -> DatasetTask:
    return DatasetTask(
        id=task_id,
        instruction="instruction",
        spreadsheet_path=f"spreadsheet/{task_id}",
        init_xlsx=Path(f"/tmp/{task_id}.xlsx"),
        metadata={},
    )


def test_cli_has_no_model_or_temperature_override() -> None:
    with pytest.raises(SystemExit):
        parse_args(
            [
                "--dataset-dir=/data",
                "--out-dir=/out",
                "--model=another-model",
            ]
        )
    with pytest.raises(SystemExit):
        parse_args(
            [
                "--dataset-dir=/data",
                "--out-dir=/out",
                "--temperature=1",
            ]
        )


def test_task_selection_preserves_manifest_order() -> None:
    tasks = [_task("a"), _task("b"), _task("c")]
    assert [task.id for task in _select_tasks(tasks, "c,a")] == ["a", "c"]


def test_task_selection_rejects_unknown_and_duplicate_ids() -> None:
    tasks = [_task("a"), _task("b")]
    with pytest.raises(ArtifactError, match="unknown"):
        _select_tasks(tasks, "missing")
    with pytest.raises(ArtifactError, match="duplicate"):
        _select_tasks(tasks, "a,a")


def test_output_must_be_outside_dataset() -> None:
    with pytest.raises(ArtifactError):
        _ensure_output_outside_dataset(Path("/data"), Path("/data/out"))
    _ensure_output_outside_dataset(Path("/data"), Path("/out"))


def test_concurrency_is_bounded() -> None:
    with pytest.raises(SystemExit):
        parse_args(["--dataset-dir=/data", "--out-dir=/out", "--concurrency=0"])
    args = parse_args(["--dataset-dir=/data", "--out-dir=/out", "--concurrency=64"])
    assert isinstance(args, argparse.Namespace)
    assert args.concurrency == 64


def test_cli_and_capture_default_to_four_workers() -> None:
    argv = ["--dataset-dir=/data", "--out-dir=/out"]

    assert parse_args(argv).concurrency == 4
    assert capture.parse_args(argv).concurrency == 4


def test_retry_failures_requires_explicit_resume() -> None:
    argv = ["--dataset-dir=/data", "--out-dir=/out", "--retry-failures"]

    with pytest.raises(SystemExit):
        parse_args(argv)
    with pytest.raises(SystemExit):
        capture.parse_args(argv)

    resumed = [*argv[:-1], "--resume", "--retry-failures"]
    assert parse_args(resumed).retry_failures is True
    assert capture.parse_args(resumed).retry_failures is True


def test_replay_write_failures_requires_resume_and_excludes_paid_retry() -> None:
    base = ["--dataset-dir=/data", "--out-dir=/out"]

    with pytest.raises(SystemExit):
        parse_args([*base, "--replay-write-failures"])
    with pytest.raises(SystemExit):
        capture.parse_args([*base, "--replay-write-failures"])
    with pytest.raises(SystemExit):
        parse_args([*base, "--resume", "--retry-failures", "--replay-write-failures"])
    with pytest.raises(SystemExit):
        capture.parse_args([*base, "--resume", "--retry-failures", "--replay-write-failures"])

    resumed = [*base, "--resume", "--replay-write-failures"]
    assert parse_args(resumed).replay_write_failures is True
    assert capture.parse_args(resumed).replay_write_failures is True


def test_preflight_describes_pinned_native_transport(
    capsys: pytest.CaptureFixture[str],
) -> None:
    tasks = [_task("a"), _task("b")]

    cli._preflight(tasks, tasks[:1])

    assert capsys.readouterr().out == (
        "preflight ok  tasks=2  selected=1  model=Qwen/Qwen3.8-27B  "
        "transport=tinker-native-sampling/0.27.1  "
        "renderer=transformers-chat-template/qwen3.8-disable-thinking-pinned  "
        "thinking=disabled  "
        "tool_enforcement=rendered_prompt_plus_strict_parser:submit_spreadsheet_answer\n"
    )


def _two_task_dataset(tmp_path: Path) -> Path:
    dataset = tmp_path / "dataset"
    records = []
    for index, task_id in enumerate(("task-1", "task-2"), start=1):
        task_dir = dataset / "spreadsheet" / task_id
        task_dir.mkdir(parents=True)
        workbook = openpyxl.Workbook()
        workbook.active.title = "Inputs"
        workbook.active["A1"] = index
        workbook.save(task_dir / f"{index}_{task_id}_init.xlsx")
        workbook.close()
        records.append(
            {
                "id": task_id,
                "instruction": "Copy A1 to B1.",
                "spreadsheet_path": f"spreadsheet/{task_id}",
                "answer_position": "B1",
                "answer_sheet": "Inputs",
                "data_position": "Inputs!A1",
            }
        )
    (dataset / "dataset.json").write_text(json.dumps(records), encoding="utf-8")
    return dataset


def test_fresh_direct_cli_creates_and_holds_its_own_output_lock(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    dataset = _two_task_dataset(tmp_path)
    out = tmp_path / "fresh-out"

    class DirectProvider:
        async def complete(self, *, system_prompt: str, user_prompt: str) -> Completion:
            del system_prompt, user_prompt
            return Completion(
                text='{"cells":[{"sheet":"Inputs","cell":"B1","value":"=A1"}]}',
                input_tokens=10,
                output_tokens=5,
                model=MODEL_ID,
                renderer="fake-test-renderer",
            )

        async def close(self) -> None:
            return None

    monkeypatch.delenv(SUPERVISOR_LOCK_FD_ENV, raising=False)
    monkeypatch.setattr(cli, "require_provider_key", lambda: None)
    monkeypatch.setattr(cli, "TinkerProvider", DirectProvider)
    args = parse_args(
        [
            f"--dataset-dir={dataset}",
            f"--out-dir={out}",
            "--ids=task-1",
            "--concurrency=1",
        ]
    )

    assert asyncio.run(cli._run(args)) == 0
    assert (out / "run.log").is_file()


def test_plain_resume_skips_failure_and_explicit_retry_can_replace_it(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    dataset = _two_task_dataset(tmp_path)
    out = tmp_path / "out"
    out.mkdir()
    (out / "run.log").write_bytes(b"captured\n")

    class AlwaysFailProvider:
        async def complete(self, *, system_prompt: str, user_prompt: str) -> object:
            raise RuntimeError("private provider details")

        async def close(self) -> None:
            return None

    monkeypatch.setattr(cli, "require_provider_key", lambda: None)
    monkeypatch.setattr(cli, "TinkerProvider", AlwaysFailProvider)
    first_args = parse_args(
        [
            f"--dataset-dir={dataset}",
            f"--out-dir={out}",
            "--ids=task-1",
            "--concurrency=1",
        ]
    )

    assert asyncio.run(cli._run(first_args)) == 1

    def key_must_not_be_requested() -> None:
        raise AssertionError("a fully checkpointed resume must not request a provider key")

    monkeypatch.setattr(cli, "require_provider_key", key_must_not_be_requested)
    resume_args = parse_args(
        [
            f"--dataset-dir={dataset}",
            f"--out-dir={out}",
            "--ids=task-1",
            "--resume",
        ]
    )
    assert resume_args.resume is True
    assert asyncio.run(cli._run(resume_args)) == 1

    key_checks = 0

    def key_is_required_for_explicit_retry() -> None:
        nonlocal key_checks
        key_checks += 1

    class RecoveringProvider:
        calls = 0

        async def complete(self, *, system_prompt: str, user_prompt: str) -> Completion:
            del system_prompt, user_prompt
            type(self).calls += 1
            return Completion(
                text='{"cells":[{"sheet":"Inputs","cell":"B1","value":"=A1"}]}',
                input_tokens=10,
                output_tokens=5,
                model=MODEL_ID,
                renderer="fake-test-renderer",
            )

        async def close(self) -> None:
            return None

    monkeypatch.setattr(cli, "require_provider_key", key_is_required_for_explicit_retry)
    monkeypatch.setattr(cli, "TinkerProvider", RecoveringProvider)
    retry_args = parse_args(
        [
            f"--dataset-dir={dataset}",
            f"--out-dir={out}",
            "--ids=task-1",
            "--resume",
            "--retry-failures",
        ]
    )

    assert asyncio.run(cli._run(retry_args)) == 0
    assert key_checks == 1
    assert RecoveringProvider.calls == 1
    prediction = json.loads((out / "predictions.jsonl").read_text(encoding="utf-8"))
    assert prediction["status"] == "ok"

    def provider_must_not_be_constructed() -> object:
        raise AssertionError("no validated failure remains to retry")

    monkeypatch.setattr(cli, "require_provider_key", key_must_not_be_requested)
    monkeypatch.setattr(cli, "TinkerProvider", provider_must_not_be_constructed)
    assert asyncio.run(cli._run(retry_args)) == 0


def test_complete_replay_path_never_requests_or_constructs_provider(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    dataset = _two_task_dataset(tmp_path)
    out = tmp_path / "out"
    out.mkdir()
    (out / "run.log").write_bytes(b"captured\n")

    class AlwaysFailProvider:
        async def complete(self, *, system_prompt: str, user_prompt: str) -> object:
            del system_prompt, user_prompt
            raise RuntimeError("private provider details")

        async def close(self) -> None:
            return None

    monkeypatch.setattr(cli, "require_provider_key", lambda: None)
    monkeypatch.setattr(cli, "TinkerProvider", AlwaysFailProvider)
    initial = parse_args(
        [
            f"--dataset-dir={dataset}",
            f"--out-dir={out}",
            "--ids=task-1",
        ]
    )
    assert asyncio.run(cli._run(initial)) == 1

    replay_calls = 0

    def local_replay(**kwargs: object) -> SimpleNamespace:
        nonlocal replay_calls
        assert kwargs["expected_model"] == MODEL_ID
        replay_calls += 1
        return SimpleNamespace(
            eligible_tasks=0,
            replayed_tasks=(),
            still_failed=(),
            model_calls=0,
        )

    def provider_access_forbidden() -> object:
        raise AssertionError("stored-response replay must not access the provider")

    monkeypatch.setattr(cli, "replay_stored_responses", local_replay)
    monkeypatch.setattr(cli, "require_provider_key", provider_access_forbidden)
    monkeypatch.setattr(cli, "TinkerProvider", provider_access_forbidden)
    replay_args = parse_args(
        [
            f"--dataset-dir={dataset}",
            f"--out-dir={out}",
            "--ids=task-1",
            "--resume",
            "--replay-write-failures",
        ]
    )

    assert asyncio.run(cli._run(replay_args)) == 1
    assert replay_calls == 1


def test_capture_resume_flag_is_forwarded_to_child_and_log_supervisor(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    dataset = _two_task_dataset(tmp_path)
    out = tmp_path / "out"
    out.mkdir()
    run_log = out / "run.log"
    run_log.write_bytes(b"prior bytes\n")
    observed: dict[str, object] = {}

    def fake_run_captured(
        command: list[str], output: Path, *, resume: bool = False
    ) -> CaptureResult:
        observed["command"] = tuple(command)
        observed["output"] = output
        observed["resume"] = resume
        return CaptureResult(tuple(command), 0, run_log)

    monkeypatch.setattr(capture, "run_captured", fake_run_captured)

    returncode = capture.main(
        [
            f"--dataset-dir={dataset}",
            f"--out-dir={out}",
            "--ids=task-2,task-1",
            "--resume",
            "--retry-failures",
        ]
    )

    assert returncode == 0
    assert observed["resume"] is True
    command = observed["command"]
    assert isinstance(command, tuple)
    assert "--resume" in command
    assert "--retry-failures" in command
    assert "--concurrency=4" in command
    # Capture preserves the caller's ID order; the child applies the existing
    # manifest-order selection semantics.
    assert "--ids=task-2,task-1" in command


def test_capture_replay_flag_is_forwarded(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    dataset = _two_task_dataset(tmp_path)
    out = tmp_path / "out"
    out.mkdir()
    run_log = out / "run.log"
    run_log.write_bytes(b"prior bytes\n")
    observed: dict[str, object] = {}

    def fake_run_captured(
        command: list[str], output: Path, *, resume: bool = False
    ) -> CaptureResult:
        observed["command"] = tuple(command)
        observed["output"] = output
        observed["resume"] = resume
        return CaptureResult(tuple(command), 0, run_log)

    monkeypatch.setattr(capture, "run_captured", fake_run_captured)

    returncode = capture.main(
        [
            f"--dataset-dir={dataset}",
            f"--out-dir={out}",
            "--resume",
            "--replay-write-failures",
        ]
    )

    assert returncode == 0
    assert observed["resume"] is True
    command = observed["command"]
    assert isinstance(command, tuple)
    assert "--resume" in command
    assert "--replay-write-failures" in command
