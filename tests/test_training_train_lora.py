from __future__ import annotations

import hashlib
import json
import os
from dataclasses import replace
from pathlib import Path
from types import SimpleNamespace
from typing import Any

import pytest

from formulabench.constants import TINKER_TELEMETRY_ENV_VAR, TINKER_TELEMETRY_VALUE
from formulabench.prompts import SYSTEM_PROMPT
from formulabench.provider import ANSWER_TOOL_NAME
from training.train_lora import (
    CORPUS_BUILDER,
    MAX_TARGET_CELLS,
    PARTITION_SEED,
    PreparedCorpus,
    PreparedExample,
    TrainConfig,
    TrainingExecutionError,
    TrainingInputError,
    _canonical_json,
    build_plan,
    execute_training,
    load_validated_corpus,
    parse_args,
    prepare_corpus,
)

_DIGEST_A = "a" * 64
_DIGEST_B = "b" * 64
_DIGEST_C = "c" * 64
_DIGEST_D = "d" * 64


def _sha256_bytes(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def _sha256_text(value: str) -> str:
    return _sha256_bytes(value.encode("utf-8"))


def _write_fixture(tmp_path: Path) -> tuple[Path, Path]:
    development_ids = [f"task-{index:02d}" for index in range(80)]
    dataset_digest = _DIGEST_A
    split = {
        "dataset_manifest_sha256": dataset_digest,
        "development_ids": development_ids,
        "initial_workbook_binding": {"aggregate_sha256": _DIGEST_B},
    }
    split_path = tmp_path / "public_split.json"
    split_raw = (json.dumps(split, sort_keys=True, indent=2) + "\n").encode("utf-8")
    split_path.write_bytes(split_raw)
    split_digest = _sha256_bytes(split_raw)

    rows: list[dict[str, Any]] = []
    examples: list[dict[str, Any]] = []
    for index, partition in enumerate(("train", "train", "validation")):
        task_id = development_ids[index]
        user_prompt = f"Complete the one-cell workbook task {task_id}."
        answer = _canonical_json(
            {
                "cells": [{"cell": "A1", "sheet": "Sheet1", "value": f"={index + 1}"}],
                "fills": [],
                "preserve_ranges": [],
                "unmerge_ranges": [],
            }
        )
        provenance = {
            "answer_sha256": _sha256_text(answer),
            "dataset_manifest_sha256": dataset_digest,
            "golden_workbook_sha256": _DIGEST_C,
            "initial_workbook_sha256": _DIGEST_D,
            "prompt_sha256": _sha256_text(user_prompt),
            "split_manifest_sha256": split_digest,
        }
        row = {
            "answer_cell_count": 1,
            "messages": [
                {"content": SYSTEM_PROMPT, "role": "system"},
                {"content": user_prompt, "role": "user"},
                {"content": answer, "role": "assistant"},
            ],
            "partition": partition,
            "provenance": provenance,
            "schema_version": 1,
            "split": "development",
            "target_cell_count": 1,
            "task_id": task_id,
        }
        rows.append(row)
        examples.append(
            {
                "answer_cell_count": 1,
                "answer_sha256": provenance["answer_sha256"],
                "golden_workbook_sha256": _DIGEST_C,
                "initial_workbook_sha256": _DIGEST_D,
                "partition": partition,
                "prompt_sha256": provenance["prompt_sha256"],
                "row_sha256": _sha256_text(_canonical_json(row)),
                "target_cell_count": 1,
                "task_id": task_id,
            }
        )

    corpus_raw = "".join(f"{_canonical_json(row)}\n" for row in rows).encode("utf-8")
    corpus_path = tmp_path / "corpus.jsonl"
    corpus_path.write_bytes(corpus_raw)
    exclusions = [
        {
            "reason": "target_cell_limit_exceeded",
            "target_cell_count": MAX_TARGET_CELLS + 1,
            "task_id": task_id,
        }
        for task_id in development_ids[3:]
    ]
    manifest = {
        "builder": CORPUS_BUILDER,
        "corpus_file": corpus_path.name,
        "corpus_sha256": _sha256_bytes(corpus_raw),
        "dataset_manifest_sha256": dataset_digest,
        "development_golden_binding_sha256": _DIGEST_C,
        "eligible_development_count": 3,
        "example_count": 3,
        "examples": examples,
        "excluded_development_count": 77,
        "exclusions": exclusions,
        "initial_workbook_binding_sha256": _DIGEST_B,
        "max_target_cells": MAX_TARGET_CELLS,
        "ordered_task_ids": development_ids[:3],
        "partition_counts": {"train": 2, "validation": 1},
        "partition_method": "test fixture with deterministic stratification",
        "partition_seed": PARTITION_SEED,
        "schema_version": 1,
        "split": "development",
        "split_manifest_sha256": split_digest,
        "system_prompt_sha256": _sha256_text(SYSTEM_PROMPT),
        "verified": True,
    }
    manifest_path = tmp_path / "corpus_manifest.json"
    manifest_path.write_text(
        json.dumps(manifest, indent=2, sort_keys=True) + "\n", encoding="utf-8"
    )
    return manifest_path, split_path


def _config(tmp_path: Path, *, execute: bool = False) -> TrainConfig:
    manifest_path, split_path = _write_fixture(tmp_path)
    return TrainConfig(
        manifest_path=manifest_path,
        split_manifest_path=split_path,
        execute=execute,
        project_id="test-project" if execute else None,
        run_dir=tmp_path / "run" if execute else None,
        rank=32,
        learning_rate=1e-4,
        batch_size=1,
        epochs=1,
        max_steps=2,
        max_sequence_tokens=1024,
        seed=2026,
        beta1=0.9,
        beta2=0.95,
        eps=1e-8,
        weight_decay=0.0,
        grad_clip_norm=0.0,
        state_ttl_seconds=604800,
        checkpoint_prefix="test-sft",
        allow_tokenizer_download=False,
    )


def test_validated_corpus_accepts_only_bound_development_rows(tmp_path: Path) -> None:
    config = _config(tmp_path)

    corpus = load_validated_corpus(config.manifest_path, config.split_manifest_path)

    assert [row.task_id for row in corpus.train_rows] == ["task-00", "task-01"]
    assert [row.task_id for row in corpus.validation_rows] == ["task-02"]
    assert corpus.excluded_development_count == 77


@pytest.mark.parametrize(
    ("field", "value", "message"),
    [
        ("verified", False, "not verified development-only"),
        ("split", "held_out", "not verified development-only"),
        ("system_prompt_sha256", _DIGEST_A, "production system prompt"),
    ],
)
def test_validated_corpus_refuses_untrusted_manifests(
    tmp_path: Path, field: str, value: object, message: str
) -> None:
    config = _config(tmp_path)
    manifest = json.loads(config.manifest_path.read_text(encoding="utf-8"))
    manifest[field] = value
    config.manifest_path.write_text(
        json.dumps(manifest, indent=2, sort_keys=True) + "\n", encoding="utf-8"
    )

    with pytest.raises(TrainingInputError, match=message):
        load_validated_corpus(config.manifest_path, config.split_manifest_path)


def test_validated_corpus_refuses_corpus_tampering(tmp_path: Path) -> None:
    config = _config(tmp_path)
    corpus_path = config.manifest_path.parent / "corpus.jsonl"
    corpus_path.write_bytes(corpus_path.read_bytes().replace(b"=1", b"=9", 1))

    with pytest.raises(TrainingInputError, match="corpus digest"):
        load_validated_corpus(config.manifest_path, config.split_manifest_path)


class _FakePrompt:
    def __init__(self, tokens: list[int]) -> None:
        self._tokens = tokens
        self.length = len(tokens)

    def to_ints(self) -> list[int]:
        return self._tokens


class _FakeRenderer:
    def __init__(self) -> None:
        self.seen_tools: list[object] = []

    def create_conversation_prefix_with_tools(
        self, tools: list[object], *, system_prompt: str
    ) -> list[dict[str, str]]:
        self.seen_tools = tools
        return [{"role": "system", "content": system_prompt}]

    def build_generation_prompt(self, messages: list[dict[str, str]]) -> _FakePrompt:
        assert messages[-1]["role"] == "user"
        return _FakePrompt([10, 11, 12])


class _FakeTokenizer:
    def apply_chat_template(self, *_args: object, **kwargs: object) -> list[int]:
        assert kwargs["add_generation_prompt"] is True
        assert kwargs["enable_thinking"] is False
        assert kwargs["tokenize"] is True
        return [10, 11, 12]


def _fake_dependencies(
    *,
    datum_length: int = 8,
) -> tuple[tuple[Any, Any, Any, Any], list[list[dict[str, Any]]]]:
    renderer = _FakeRenderer()
    converted: list[list[dict[str, Any]]] = []

    def convert(messages: list[dict[str, Any]], _renderer: object) -> Any:
        converted.append(messages)
        weights = [0.0, 0.0, 1.0, *([1.0] * (datum_length - 3))]
        return SimpleNamespace(
            model_input=SimpleNamespace(length=datum_length),
            loss_fn_inputs={"weights": SimpleNamespace(data=weights)},
        )

    def tool_call(answer: str) -> Any:
        return SimpleNamespace(
            function=SimpleNamespace(name=ANSWER_TOOL_NAME, arguments=answer), id=None
        )

    return (_FakeTokenizer(), renderer, convert, tool_call), converted


def test_preflight_uses_production_tool_wire_format_and_accounts_for_validation(
    tmp_path: Path,
) -> None:
    config = _config(tmp_path)
    corpus = load_validated_corpus(config.manifest_path, config.split_manifest_path)
    dependencies, converted = _fake_dependencies()

    prepared = prepare_corpus(corpus, max_sequence_tokens=1024, dependencies=dependencies)
    plan = build_plan(config, prepared)

    assert prepared.wire_format_match is True
    assert dependencies[1].seen_tools[0]["name"] == ANSWER_TOOL_NAME
    assistant = converted[0][-1]
    assert assistant["content"] == ""
    assert assistant["tool_calls"][0].function.name == ANSWER_TOOL_NAME
    assert assistant["tool_calls"][0].function.arguments == corpus.rows[0].messages[-1]["content"]
    assert plan["provider_calls"] == "zero"
    assert plan["train"] == {
        "assistant_completion_tokens": 12,
        "examples": 2,
        "max_sequence_tokens": 8,
        "total_tokens": 16,
    }
    assert plan["validation"] == {
        "assistant_completion_tokens": 6,
        "examples": 1,
        "max_sequence_tokens": 8,
        "total_tokens": 8,
    }


def test_preflight_refuses_a_later_row_wire_format_mismatch(tmp_path: Path) -> None:
    config = _config(tmp_path)
    corpus = load_validated_corpus(config.manifest_path, config.split_manifest_path)
    dependencies, _converted = _fake_dependencies()

    class LaterMismatchTokenizer(_FakeTokenizer):
        def __init__(self) -> None:
            self.calls = 0

        def apply_chat_template(self, *_args: object, **kwargs: object) -> list[int]:
            tokens = super().apply_chat_template(*_args, **kwargs)
            self.calls += 1
            return tokens if self.calls == 1 else [10, 11, 99]

    tokenizer = LaterMismatchTokenizer()
    dependencies = (tokenizer, dependencies[1], dependencies[2], dependencies[3])

    with pytest.raises(TrainingInputError, match="production wire format for task-01"):
        prepare_corpus(corpus, max_sequence_tokens=1024, dependencies=dependencies)

    assert tokenizer.calls == 2


def test_preflight_refuses_overlength_instead_of_truncating(tmp_path: Path) -> None:
    config = _config(tmp_path)
    corpus = load_validated_corpus(config.manifest_path, config.split_manifest_path)
    dependencies, _converted = _fake_dependencies(datum_length=1025)

    with pytest.raises(TrainingInputError, match="refusing truncation"):
        prepare_corpus(corpus, max_sequence_tokens=1024, dependencies=dependencies)


class _Future:
    def __init__(self, events: list[str], name: str, result: object = None) -> None:
        self.events = events
        self.name = name
        self.value = result

    def result(self) -> object:
        self.events.append(f"wait:{self.name}")
        return self.value


class _Datum:
    def __init__(self, label: str) -> None:
        self.label = label
        self.loss_fn_inputs = {"weights": SimpleNamespace(data=[1.0])}


def _loss_result(logprob: float, *, metric: float) -> Any:
    return SimpleNamespace(
        loss_fn_outputs=[{"logprobs": SimpleNamespace(data=[logprob])}],
        metrics={"tokens_per_second": metric, "unsafe metric name": "not numeric"},
    )


class _TrainingClient:
    def __init__(self, events: list[str]) -> None:
        self.events = events
        self.training_data: list[list[object]] = []

    def forward(self, data: list[object], loss: str) -> _Future:
        assert loss == "cross_entropy"
        assert [item.label for item in data] == ["validation"]
        count = sum(event.startswith("submit:validation") for event in self.events) + 1
        name = f"validation-{count}"
        self.events.append(f"submit:{name}")
        return _Future(
            self.events,
            name,
            _loss_result(-2.0 if count == 1 else -1.0, metric=float(count)),
        )

    def forward_backward(self, data: list[object], loss: str) -> _Future:
        assert loss == "cross_entropy"
        self.training_data.append(data)
        name = f"forward-{len(self.training_data)}"
        self.events.append(f"submit:{name}")
        return _Future(self.events, name, _loss_result(-1.5, metric=12.0))

    def optim_step(self, _adam: object) -> _Future:
        name = f"optim-{sum(event.startswith('submit:optim') for event in self.events) + 1}"
        self.events.append(f"submit:{name}")
        result = SimpleNamespace(metrics={"learning_rate": 1e-4})
        return _Future(self.events, name, result)

    def save_state(self, name: str, **kwargs: object) -> _Future:
        assert kwargs["ttl_seconds"] == 604800
        assert kwargs["overwrite"] is False
        self.events.append("submit:save-state")
        return _Future(self.events, "save-state", SimpleNamespace(path=f"tinker://{name}"))

    def save_weights_for_sampler(self, name: str, **kwargs: object) -> _Future:
        assert kwargs["ttl_seconds"] is None
        self.events.append("submit:save-sampler")
        return _Future(self.events, "save-sampler", SimpleNamespace(path=f"tinker://{name}"))


class _Service:
    def __init__(self, training_client: _TrainingClient) -> None:
        self.training_client = training_client

    def create_lora_training_client(self, **kwargs: object) -> _TrainingClient:
        assert kwargs["base_model"] == "Qwen/Qwen3.8-27B"
        assert kwargs["rank"] == 32
        return self.training_client


def test_execute_pipelines_train_only_and_persists_sampler_checkpoint(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    config = _config(tmp_path, execute=True)
    corpus = load_validated_corpus(config.manifest_path, config.split_manifest_path)
    prepared = PreparedCorpus(
        corpus=corpus,
        examples=(
            PreparedExample("task-00", "train", _Datum("train-a"), 8, 4),
            PreparedExample("task-01", "train", _Datum("train-b"), 8, 4),
            PreparedExample("task-02", "validation", _Datum("validation"), 8, 4),
        ),
        wire_format_match=True,
    )
    plan = build_plan(config, prepared)
    events: list[str] = []
    client = _TrainingClient(events)
    service = _Service(client)
    service_calls: list[dict[str, object]] = []
    monkeypatch.setenv(TINKER_TELEMETRY_ENV_VAR, "1")

    def service_factory(**kwargs: object) -> _Service:
        assert os.environ[TINKER_TELEMETRY_ENV_VAR] == TINKER_TELEMETRY_VALUE
        service_calls.append(kwargs)
        return service

    checkpoints = execute_training(
        config,
        prepared,
        plan,
        service_factory=service_factory,
        adam_factory=lambda **kwargs: kwargs,
        nll_calculator=lambda logprobs, _weights: (
            -sum(item.data[0] for item in logprobs) / len(logprobs)
        ),
        environ={"TINKER_API_KEY": "not-recorded"},
    )

    assert service_calls == [
        {
            "project_id": "test-project",
            "user_metadata": {"corpus_sha256": corpus.corpus_sha256, "recipe": CORPUS_BUILDER},
        }
    ]
    assert sorted(item.label for batch in client.training_data for item in batch) == [
        "train-a",
        "train-b",
    ]
    assert "validation" not in [item.label for batch in client.training_data for item in batch]
    assert events.index("submit:validation-1") < events.index("submit:forward-1")
    assert events.index("submit:forward-2") < events.index("wait:forward-1")
    assert events.index("wait:optim-2") < events.index("submit:validation-2")
    assert events.index("wait:validation-2") < events.index("submit:save-state")
    assert checkpoints["sampler"]["path"] == "tinker://test-sft-sampler"
    metadata_text = (config.run_dir / "run_metadata.json").read_text(encoding="utf-8")
    metadata = json.loads(metadata_text)
    assert '"status": "complete"' in metadata_text
    assert "TINKER_API_KEY" not in metadata_text
    assert metadata["validation_nll_before"] == 2.0
    assert metadata["validation_nll_after"] == 1.0
    assert metadata["validation_nll_change"] == -1.0
    assert [step["train_nll"] for step in metadata["training_steps"]] == [1.5, 1.5]
    assert metadata["training_steps"][0]["forward_metrics"] == {"tokens_per_second": 12.0}


def test_paid_training_requires_explicit_execute_before_service_creation(tmp_path: Path) -> None:
    config = _config(tmp_path)
    corpus = load_validated_corpus(config.manifest_path, config.split_manifest_path)
    prepared = PreparedCorpus(corpus=corpus, examples=(), wire_format_match=True)
    called = False

    def service_factory(**_kwargs: object) -> None:
        nonlocal called
        called = True

    with pytest.raises(TrainingInputError, match="requires --execute"):
        execute_training(
            config,
            prepared,
            {},
            service_factory=service_factory,
            adam_factory=lambda **kwargs: kwargs,
        )
    assert called is False


def test_paid_training_requires_api_key_before_service_creation(tmp_path: Path) -> None:
    config = _config(tmp_path, execute=True)
    corpus = load_validated_corpus(config.manifest_path, config.split_manifest_path)
    prepared = PreparedCorpus(
        corpus=corpus,
        examples=(
            PreparedExample("task-00", "train", _Datum("train-a"), 8, 4),
            PreparedExample("task-02", "validation", _Datum("validation"), 8, 4),
        ),
        wire_format_match=True,
    )
    called = False

    def service_factory(**_kwargs: object) -> None:
        nonlocal called
        called = True

    with pytest.raises(TrainingInputError, match="TINKER_API_KEY is required"):
        execute_training(
            config,
            prepared,
            build_plan(config, prepared),
            service_factory=service_factory,
            adam_factory=lambda **kwargs: kwargs,
            nll_calculator=lambda _logprobs, _weights: 1.0,
            environ={},
        )
    assert called is False


def test_paid_training_refuses_reused_run_directory_before_service_creation(
    tmp_path: Path,
) -> None:
    config = _config(tmp_path, execute=True)
    assert config.run_dir is not None
    config.run_dir.mkdir()
    retained = config.run_dir / "run_metadata.json"
    retained.write_text('{"status":"complete"}\n', encoding="utf-8")
    corpus = load_validated_corpus(config.manifest_path, config.split_manifest_path)
    prepared = PreparedCorpus(
        corpus=corpus,
        examples=(
            PreparedExample("task-00", "train", _Datum("train-a"), 8, 4),
            PreparedExample("task-02", "validation", _Datum("validation"), 8, 4),
        ),
        wire_format_match=True,
    )
    called = False

    def service_factory(**_kwargs: object) -> None:
        nonlocal called
        called = True

    with pytest.raises(TrainingInputError, match="new path"):
        execute_training(
            config,
            prepared,
            build_plan(config, prepared),
            service_factory=service_factory,
            adam_factory=lambda **kwargs: kwargs,
            nll_calculator=lambda _logprobs, _weights: 1.0,
            environ={"TINKER_API_KEY": "not-recorded"},
        )
    assert called is False
    assert retained.read_text(encoding="utf-8") == '{"status":"complete"}\n'


def test_cli_defaults_to_dry_run_and_bounds_paid_parameters() -> None:
    config = parse_args([])

    assert config.execute is False
    assert config.project_id is None
    assert config.run_dir is None
    assert config.max_sequence_tokens == 65536

    with pytest.raises(SystemExit):
        parse_args(["--max-sequence-tokens", "1000000"])
    with pytest.raises(SystemExit):
        parse_args(["--execute", "--project-id", "test-project"])


def test_provider_failure_records_only_exception_class(tmp_path: Path) -> None:
    config = _config(tmp_path, execute=True)
    corpus = load_validated_corpus(config.manifest_path, config.split_manifest_path)
    prepared = PreparedCorpus(
        corpus=corpus,
        examples=(
            PreparedExample("task-00", "train", _Datum("train-a"), 8, 4),
            PreparedExample("task-02", "validation", _Datum("validation"), 8, 4),
        ),
        wire_format_match=True,
    )
    plan = build_plan(config, prepared)
    secret = "secret-provider-detail-must-not-be-written"

    def service_factory(**_kwargs: object) -> None:
        raise RuntimeError(secret)

    with pytest.raises(TrainingExecutionError, match="named session"):
        execute_training(
            config,
            prepared,
            plan,
            service_factory=service_factory,
            adam_factory=lambda **kwargs: kwargs,
            nll_calculator=lambda _logprobs, _weights: 1.0,
            environ={"TINKER_API_KEY": "not-recorded"},
        )
    metadata_text = (config.run_dir / "run_metadata.json").read_text(encoding="utf-8")
    assert secret not in metadata_text
    assert '"error_type": "RuntimeError"' in metadata_text


def test_execute_config_cannot_be_downgraded_with_replace(tmp_path: Path) -> None:
    config = _config(tmp_path, execute=True)

    with pytest.raises(TrainingInputError, match="project-id"):
        execute_training(
            replace(config, project_id=None),
            SimpleNamespace(),
            {},
            service_factory=lambda **_kwargs: None,
            adam_factory=lambda **kwargs: kwargs,
        )
