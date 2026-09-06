from __future__ import annotations

import asyncio
import hashlib
import os
import sys
import threading
from types import SimpleNamespace

import pytest

import formulabench.provider as provider_module
from formulabench.constants import (
    KEY_ENV_VAR,
    MODEL_ID,
    MODEL_PROVENANCE,
    NUM_SAMPLES,
    PROVIDER_MAX_RETRIES,
    RENDERER_ID,
    STOP_TOKEN_ID,
    TEMPERATURE,
    THINKING_MODE,
    TOKENIZER_REVISION,
    TRANSPORT_ID,
)
from formulabench.provider import (
    ANSWER_TOOL,
    ANSWER_TOOL_NAME,
    QWEN_ANSWER_TOOL,
    ProviderConfigurationError,
    ProviderResponseError,
    TinkerProvider,
    parse_native_answer,
    require_provider_key,
    validate_sampler_checkpoint,
)


class FakeTokenizer:
    def __init__(self, decoded: str) -> None:
        self.decoded = decoded
        self.render_kwargs: dict[str, object] | None = None
        self.decode_args: tuple[list[int], dict[str, object]] | None = None

    def apply_chat_template(self, messages: object, **kwargs: object) -> list[int]:
        self.render_kwargs = {"messages": messages, **kwargs}
        return [101, 102, 103]

    def decode(self, tokens: list[int], **kwargs: object) -> str:
        self.decode_args = (tokens, kwargs)
        return self.decoded


class FakeSampler:
    def __init__(self, *, decoded_tokens: list[int], stop_reason: str = "stop") -> None:
        self.response = SimpleNamespace(
            sequences=[SimpleNamespace(tokens=decoded_tokens, stop_reason=stop_reason)]
        )
        self.calls: list[dict[str, object]] = []

    async def sample_async(self, **kwargs: object) -> object:
        self.calls.append(kwargs)
        return self.response


class FakeTinkerLoopHolder:
    def __init__(self) -> None:
        self.ready = threading.Event()
        self.loop: asyncio.AbstractEventLoop | None = None
        self.scheduled = 0
        self.close_calls = 0
        self.cleanup_calls = 0
        self.thread = threading.Thread(target=self._run, daemon=True)
        self.thread.start()
        assert self.ready.wait(timeout=2)

    def _run(self) -> None:
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        self.loop = loop
        self.ready.set()
        loop.run_forever()
        loop.close()

    def run_coroutine_threadsafe(self, coroutine: object) -> object:
        assert self.loop is not None
        self.scheduled += 1
        return asyncio.run_coroutine_threadsafe(coroutine, self.loop)

    def close(self) -> None:
        self.close_calls += 1

    async def _async_cleanup(self) -> None:
        assert asyncio.get_running_loop() is self.loop
        self.cleanup_calls += 1

    def stop(self) -> None:
        assert self.loop is not None
        self.loop.call_soon_threadsafe(self.loop.stop)
        self.thread.join(timeout=2)
        assert not self.thread.is_alive()


class FakeFuturesPoller:
    def __init__(self, holder: FakeTinkerLoopHolder) -> None:
        self.holder = holder
        self._task: asyncio.Task[None] | None = None
        self.closed_tasks: list[asyncio.Task[None]] = []
        self.drained = threading.Event()

    async def ensure_running(self) -> None:
        if self._task is None or self._task.done():
            self.drained.clear()

            async def wait_forever() -> None:
                try:
                    await asyncio.Event().wait()
                finally:
                    self.drained.set()

            self._task = asyncio.create_task(wait_forever())
            await asyncio.sleep(0)

    def close(self) -> None:
        assert asyncio.get_running_loop() is self.holder.loop
        assert self._task is not None
        self.closed_tasks.append(self._task)
        self._task.cancel()
        self._task = None


class FakeSharedNativeSampler(FakeSampler):
    def __init__(self, holder: FakeTinkerLoopHolder) -> None:
        super().__init__(decoded_tokens=[201, STOP_TOKEN_ID])
        self.holder = holder
        self._futures_poller = FakeFuturesPoller(holder)
        self.close_calls = 0

    async def sample_async(self, **kwargs: object) -> object:
        start = self.holder.run_coroutine_threadsafe(self._futures_poller.ensure_running())
        await asyncio.to_thread(start.result)
        return await super().sample_async(**kwargs)

    def close(self) -> None:
        self.close_calls += 1


def test_submission_constants_are_fixed() -> None:
    assert MODEL_ID == "Qwen/Qwen3.8-27B"
    assert KEY_ENV_VAR == "TINKER_API_KEY"
    assert TEMPERATURE == 0.0
    assert THINKING_MODE == "disabled"
    assert TOKENIZER_REVISION == "1d4bf0f2ff6012fd82039f2fa52739d0dd7c60c0"
    assert TRANSPORT_ID == "tinker-native-sampling/0.27.1"
    assert RENDERER_ID == "transformers-chat-template/qwen3.8-disable-thinking-pinned"
    assert PROVIDER_MAX_RETRIES == 0
    assert NUM_SAMPLES == 1


def test_provider_key_boundary() -> None:
    with pytest.raises(ProviderConfigurationError, match=KEY_ENV_VAR) as exc:
        require_provider_key({})
    assert "secret-value" not in str(exc.value)
    assert require_provider_key({KEY_ENV_VAR: "secret-value"}) is None


@pytest.mark.parametrize(
    "checkpoint",
    [
        "https://example.invalid/checkpoint",
        "tinker://",
        "tinker://run-id",
        "tinker://run-id/",
        "tinker://run-id/../checkpoint",
        "tinker://run-id/weights/final",
        "tinker://run-id/not_sampler/final",
        "tinker://run-id/checkpoint?token=secret",
        "tinker://run-id/checkpoint#fragment",
        "tinker://user@run-id/checkpoint",
        "tinker://run id/checkpoint",
        "tinker://run-id/chéckpoint",
        "tinker://" + "a" * 513,
    ],
)
def test_sampler_checkpoint_rejects_unbounded_or_unsafe_paths(checkpoint: str) -> None:
    with pytest.raises(ProviderConfigurationError) as caught:
        validate_sampler_checkpoint(checkpoint)

    assert str(caught.value) == "sampler checkpoint must be a safe tinker:// path"


def test_sampler_checkpoint_accepts_bounded_tinker_paths() -> None:
    for checkpoint in (
        "tinker://run-id/sampler_weights/final",
        "tinker://01234567-89ab-cdef-0123-456789abcdef:train:0/sampler_weights/checkpoint-001",
    ):
        assert validate_sampler_checkpoint(checkpoint) == checkpoint


def test_provider_refuses_an_unsafe_checkpoint_before_native_client_creation(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv(KEY_ENV_VAR, "secret-value")

    with pytest.raises(ProviderConfigurationError, match="safe tinker"):
        TinkerProvider(
            sampler_checkpoint="tinker://run-id/checkpoint?token=secret",
            sampling_client=object(),
            tokenizer=FakeTokenizer('{"cells":[]}'),
        )


def test_provider_forces_tinker_telemetry_off_before_any_client_use(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv(KEY_ENV_VAR, "secret-value")
    monkeypatch.setenv("TINKER_TELEMETRY", "1")

    TinkerProvider(sampling_client=object(), tokenizer=FakeTokenizer('{"cells":[]}'))

    assert os.environ["TINKER_TELEMETRY"] == "0"


def test_answer_tool_is_converted_to_plain_qwen_tool_shape() -> None:
    assert {
        "name": ANSWER_TOOL_NAME,
        "description": ANSWER_TOOL["description"],
        "parameters": ANSWER_TOOL["input_schema"],
    } == QWEN_ANSWER_TOOL


def test_answer_tool_schema_supports_strict_dates_and_preserve_ranges() -> None:
    properties = ANSWER_TOOL["input_schema"]["properties"]
    cell_value = properties["cells"]["items"]["properties"]["value"]
    fill_value = properties["fills"]["items"]["properties"]["value"]

    assert cell_value == fill_value
    typed_objects = {
        branch["properties"]["type"]["const"]: branch
        for branch in cell_value["oneOf"]
        if branch.get("type") == "object"
    }
    assert set(typed_objects) == {"date", "datetime"}
    assert typed_objects["date"]["additionalProperties"] is False
    assert typed_objects["date"]["required"] == ["type", "value"]
    assert typed_objects["datetime"]["additionalProperties"] is False
    assert typed_objects["datetime"]["required"] == ["type", "value"]
    assert "timezone-free" in cell_value["description"]

    preserve = properties["preserve_ranges"]
    assert preserve["type"] == "array"
    assert preserve["items"]["additionalProperties"] is False
    assert preserve["items"]["required"] == ["sheet", "range"]
    assert set(preserve["items"]["properties"]) == {"sheet", "range"}


@pytest.mark.parametrize(
    ("offline_name", "offline_value", "expects_local_only"),
    [
        (None, None, False),
        ("HF_HUB_OFFLINE", "YES", True),
        ("TRANSFORMERS_OFFLINE", "1", True),
    ],
)
def test_tokenizer_loader_fetches_pinned_revision_only_when_not_offline(
    monkeypatch: pytest.MonkeyPatch,
    offline_name: str | None,
    offline_value: str | None,
    expects_local_only: bool,
) -> None:
    calls: list[tuple[str, dict[str, object]]] = []
    sentinel = object()

    class FakeAutoTokenizer:
        @staticmethod
        def from_pretrained(model: str, **kwargs: object) -> object:
            calls.append((model, kwargs))
            return sentinel

    monkeypatch.delenv("HF_HUB_OFFLINE", raising=False)
    monkeypatch.delenv("TRANSFORMERS_OFFLINE", raising=False)
    monkeypatch.delenv("TRANSFORMERS_VERBOSITY", raising=False)
    if offline_name is not None and offline_value is not None:
        monkeypatch.setenv(offline_name, offline_value)
    monkeypatch.setitem(
        sys.modules,
        "transformers",
        SimpleNamespace(AutoTokenizer=FakeAutoTokenizer),
    )
    provider_module._load_tokenizer.cache_clear()

    try:
        assert provider_module._load_tokenizer() is sentinel
    finally:
        provider_module._load_tokenizer.cache_clear()

    expected_options: dict[str, object] = {
        "revision": TOKENIZER_REVISION,
        "trust_remote_code": False,
    }
    if expects_local_only:
        expected_options["local_files_only"] = True
    assert calls == [(MODEL_ID, expected_options)]
    assert provider_module.os.environ["TRANSFORMERS_VERBOSITY"] == "error"


def test_native_parser_accepts_anchored_tool_call_and_preserves_raw_json() -> None:
    raw_cells = '[ {"sheet":"Inputs","cell":"B2","value":"=SUM(A1:A2)"} ]'
    text = (
        "<tool_call>\n<function=submit_spreadsheet_answer>\n"
        f"<parameter=cells>\n{raw_cells}\n</parameter>\n"
        '<parameter=fills>\n[{"sheet":"Inputs","range":"C2:C3","value":"=B2"}]'
        "\n</parameter>\n"
        '<parameter=preserve_ranges>\n[{"sheet":"Inputs","range":"D2:D3"}]'
        "\n</parameter>\n</function>\n</tool_call>"
    )

    parsed, response_format = parse_native_answer(text)

    assert response_format == "qwen_xml_tool_call"
    assert raw_cells in parsed
    assert parsed.startswith('{"cells":')
    assert '"fills":[' in parsed
    assert '"preserve_ranges":[' in parsed


def test_native_parser_accepts_only_complete_strict_json_fallback() -> None:
    raw = ' {"cells":[{"sheet":"Inputs","cell":"B2","value":7}]} '
    assert parse_native_answer(raw) == (raw.strip(), "strict_json_fallback")

    for invalid in (
        'analysis {"cells":[]}',
        '```json\n{"cells":[]}\n```',
        '{"cells":[]} trailing',
        '{"cells":[],"cells":[]}',
        '{"cells":[{"value":NaN}]}',
        '{"cells":[{"value":1e999}]}',
    ):
        with pytest.raises(ProviderResponseError):
            parse_native_answer(invalid)


@pytest.mark.parametrize(
    "invalid,rejection",
    [
        (
            "<tool_call><function=other><parameter=cells>[]</parameter></function></tool_call>",
            "rejected_tool_name",
        ),
        (
            "<tool_call><function=submit_spreadsheet_answer>"
            "<parameter=cells>[]</parameter><parameter=cells>[]</parameter>"
            "</function></tool_call>",
            "rejected_duplicate_parameter",
        ),
        (
            "<tool_call><function=submit_spreadsheet_answer>"
            "<parameter=unknown>[]</parameter></function></tool_call>",
            "rejected_parameter_name",
        ),
        (
            "<tool_call><function=submit_spreadsheet_answer>"
            "<parameter=fills>[]</parameter></function></tool_call>",
            "rejected_missing_parameter",
        ),
        (
            "<tool_call><function=submit_spreadsheet_answer>"
            '<parameter=cells>[{"x":1,"x":2}]</parameter></function></tool_call>',
            "rejected_json",
        ),
        (
            "<tool_call><function=submit_spreadsheet_answer>"
            "<parameter=cells>[]</parameter></function></tool_call> dangling",
            "rejected_tool_grammar",
        ),
        (
            "<tool_call><function=submit_spreadsheet_answer>"
            "<parameter=cells>[]</parameter></function></tool_call>"
            "<tool_call><function=submit_spreadsheet_answer>"
            "<parameter=cells>[]</parameter></function></tool_call>",
            "rejected_tool_grammar",
        ),
    ],
)
def test_native_parser_rejects_malformed_or_ambiguous_calls(invalid: str, rejection: str) -> None:
    with pytest.raises(ProviderResponseError) as caught:
        parse_native_answer(invalid)
    assert caught.value.response_rejection == rejection


def test_provider_renders_and_sends_one_fixed_native_sample(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    answer = (
        "<tool_call><function=submit_spreadsheet_answer>"
        '<parameter=cells>[{"sheet":"Inputs","cell":"C3","value":7}]</parameter>'
        "</function></tool_call>"
    )
    tokenizer = FakeTokenizer(answer)
    sampler = FakeSampler(decoded_tokens=[201, 202, STOP_TOKEN_ID])
    monkeypatch.setenv(KEY_ENV_VAR, "secret-value")
    provider = TinkerProvider(sampling_client=sampler, tokenizer=tokenizer)

    completion = asyncio.run(provider.complete(system_prompt="system", user_prompt="user"))

    assert completion.text == '{"cells":[{"sheet":"Inputs","cell":"C3","value":7}]}'
    assert completion.input_tokens == 3
    assert completion.output_tokens == 3
    assert completion.model == MODEL_ID
    assert completion.model_provenance == MODEL_PROVENANCE
    assert completion.renderer == RENDERER_ID
    assert completion.stop_reason == "stop"
    assert completion.parse_termination == "stop_sequence"
    assert completion.response_format == "qwen_xml_tool_call"
    assert tokenizer.render_kwargs == {
        "messages": [
            {"role": "system", "content": "system"},
            {"role": "user", "content": "user"},
        ],
        "tools": [QWEN_ANSWER_TOOL],
        "tokenize": True,
        "add_generation_prompt": True,
        "enable_thinking": False,
        "return_dict": False,
    }
    assert tokenizer.decode_args == (
        [201, 202],
        {"skip_special_tokens": False, "clean_up_tokenization_spaces": False},
    )
    assert len(sampler.calls) == 1
    call = sampler.calls[0]
    assert call["num_samples"] == 1
    assert call["prompt"].to_ints() == [101, 102, 103]
    params = call["sampling_params"]
    assert params.max_tokens == 8192
    assert params.temperature == 0.0
    assert params.stop == [STOP_TOKEN_ID]


def test_native_service_and_sampler_are_process_scoped_with_retries_disabled(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import tinker

    service_kwargs: list[dict[str, object]] = []
    sampler_kwargs: list[dict[str, object]] = []
    telemetry_values: list[str | None] = []
    sampler = object()

    class FakeServiceClient:
        def __init__(self, **kwargs: object) -> None:
            service_kwargs.append(kwargs)
            telemetry_values.append(os.environ.get("TINKER_TELEMETRY"))

        def create_sampling_client(self, **kwargs: object) -> object:
            sampler_kwargs.append(kwargs)
            return sampler

    monkeypatch.setenv(KEY_ENV_VAR, "secret-value")
    monkeypatch.setenv("TINKER_TELEMETRY", "1")
    monkeypatch.setattr(tinker, "ServiceClient", FakeServiceClient)
    monkeypatch.setattr(provider_module, "_SHARED_SERVICE_CLIENT", None)
    monkeypatch.setattr(provider_module, "_SHARED_SAMPLING_CLIENT", None)

    assert provider_module._shared_sampling_client() is sampler
    assert provider_module._shared_sampling_client() is sampler
    assert os.environ["TINKER_TELEMETRY"] == "0"
    assert telemetry_values == ["0"]
    assert service_kwargs == [{"api_key": "secret-value", "max_retries": 0}]
    assert len(sampler_kwargs) == 1
    assert sampler_kwargs[0]["base_model"] == MODEL_ID
    assert sampler_kwargs[0]["retry_config"].enable_retry_logic is False


def test_checkpoint_sampler_is_run_local_with_hashed_provenance_and_no_retries(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import tinker

    checkpoint = "tinker://01234567-89ab-cdef-0123-456789abcdef:train:0/sampler_weights/final"
    answer = (
        "<tool_call><function=submit_spreadsheet_answer>"
        "<parameter=cells>[]</parameter></function></tool_call>"
    )
    sampler = FakeSampler(decoded_tokens=[201, STOP_TOKEN_ID])
    service_kwargs: list[dict[str, object]] = []
    sampler_kwargs: list[dict[str, object]] = []
    services: list[object] = []

    class FakeServiceClient:
        def __init__(self, **kwargs: object) -> None:
            service_kwargs.append(kwargs)
            services.append(self)

        def create_sampling_client(self, **kwargs: object) -> object:
            sampler_kwargs.append(kwargs)
            return sampler

    monkeypatch.setenv(KEY_ENV_VAR, "secret-value")
    monkeypatch.setenv("TINKER_TELEMETRY", "1")
    monkeypatch.setattr(tinker, "ServiceClient", FakeServiceClient)
    provider = TinkerProvider(
        sampler_checkpoint=checkpoint,
        tokenizer=FakeTokenizer(answer),
    )

    completion = asyncio.run(provider.complete(system_prompt="system", user_prompt="user"))

    expected_provenance = (
        "sampler_checkpoint_sha256:" + hashlib.sha256(checkpoint.encode("utf-8")).hexdigest()
    )
    assert completion.model == MODEL_ID
    assert completion.model_provenance == expected_provenance
    assert provider.model_provenance == expected_provenance
    assert checkpoint not in completion.model_provenance
    assert provider._uses_shared_sampling_client is False
    assert provider._service_client is services[0]
    assert os.environ["TINKER_TELEMETRY"] == "0"
    assert service_kwargs == [{"api_key": "secret-value", "max_retries": 0}]
    assert len(sampler_kwargs) == 1
    assert sampler_kwargs[0]["base_model"] == MODEL_ID
    assert sampler_kwargs[0]["model_path"] == checkpoint
    assert sampler_kwargs[0]["retry_config"].enable_retry_logic is False


def test_checkpoint_mode_preserves_injected_sampler_behavior(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import tinker

    checkpoint = "tinker://01234567-89ab-cdef-0123-456789abcdef:train:0/sampler_weights/final"
    answer = (
        "<tool_call><function=submit_spreadsheet_answer>"
        "<parameter=cells>[]</parameter></function></tool_call>"
    )
    sampler = FakeSampler(decoded_tokens=[201, STOP_TOKEN_ID])

    def native_client_must_not_be_created(**kwargs: object) -> object:
        del kwargs
        raise AssertionError("injected sampling clients must remain authoritative")

    monkeypatch.setenv(KEY_ENV_VAR, "secret-value")
    monkeypatch.setattr(tinker, "ServiceClient", native_client_must_not_be_created)
    provider = TinkerProvider(
        sampler_checkpoint=checkpoint,
        sampling_client=sampler,
        tokenizer=FakeTokenizer(answer),
    )

    completion = asyncio.run(provider.complete(system_prompt="system", user_prompt="user"))
    asyncio.run(provider.close())

    assert len(sampler.calls) == 1
    assert completion.model_provenance.startswith("sampler_checkpoint_sha256:")


def test_provider_close_drains_only_shared_sampler_poller_and_preserves_reuse(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv(KEY_ENV_VAR, "secret-value")
    monkeypatch.setattr(provider_module, "version", lambda package: "0.27.1")
    holder = FakeTinkerLoopHolder()
    sampler = FakeSharedNativeSampler(holder)
    service = object()
    monkeypatch.setattr(provider_module, "_SHARED_SERVICE_CLIENT", service)
    monkeypatch.setattr(provider_module, "_SHARED_SAMPLING_CLIENT", sampler)
    provider = TinkerProvider(tokenizer=FakeTokenizer('{"cells":[]}'))

    async def exercise() -> None:
        await provider.complete(system_prompt="system", user_prompt="user")
        first_task = sampler._futures_poller._task
        assert first_task is not None
        assert sampler._futures_poller.closed_tasks == []

        await provider.close()
        assert sampler._futures_poller.drained.is_set()
        assert first_task.done()
        assert sampler._futures_poller.closed_tasks == [first_task]

        # Repeated close is harmless, and the same process-scoped sampler can
        # start a fresh poller task for a later completion.
        await provider.close()
        assert sampler._futures_poller.closed_tasks == [first_task]
        await provider.complete(system_prompt="system", user_prompt="user")
        second_task = sampler._futures_poller._task
        assert second_task is not None
        assert second_task is not first_task
        await provider.close()
        assert sampler._futures_poller.drained.is_set()
        assert second_task.done()
        assert sampler._futures_poller.closed_tasks == [first_task, second_task]

    try:
        asyncio.run(exercise())
    finally:
        holder.stop()

    assert len(sampler.calls) == 2
    assert provider._sampling_client is sampler
    assert provider_module._SHARED_SERVICE_CLIENT is service
    assert provider_module._SHARED_SAMPLING_CLIENT is sampler
    assert sampler.close_calls == 0
    assert holder.close_calls == 0


def test_provider_close_drains_run_local_checkpoint_sampler_poller(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import tinker

    monkeypatch.setenv(KEY_ENV_VAR, "secret-value")
    monkeypatch.setattr(provider_module, "version", lambda package: "0.27.1")
    holder = FakeTinkerLoopHolder()
    sampler = FakeSharedNativeSampler(holder)
    services: list[object] = []

    class FakeServiceClient:
        def __init__(self, **kwargs: object) -> None:
            del kwargs
            self._session_holder = holder
            services.append(self)

        def create_sampling_client(self, **kwargs: object) -> object:
            del kwargs
            return sampler

    monkeypatch.setattr(tinker, "ServiceClient", FakeServiceClient)
    provider = TinkerProvider(
        sampler_checkpoint=(
            "tinker://01234567-89ab-cdef-0123-456789abcdef:train:0/sampler_weights/final"
        ),
        tokenizer=FakeTokenizer('{"cells":[]}'),
    )

    async def exercise() -> None:
        await provider.complete(system_prompt="system", user_prompt="user")
        first_poller_task = sampler._futures_poller._task
        assert first_poller_task is not None

        await provider.close()

        assert sampler._futures_poller.drained.is_set()
        assert first_poller_task.done()
        assert sampler._futures_poller.closed_tasks == [first_poller_task]
        assert holder.cleanup_calls == 1
        assert provider._sampling_client is None
        assert provider._service_client is None

        # Closing releases the run-local client. Reusing the provider creates
        # a new service rather than reviving a closed checkpoint session.
        await provider.complete(system_prompt="system", user_prompt="user")
        second_poller_task = sampler._futures_poller._task
        assert second_poller_task is not None
        assert second_poller_task is not first_poller_task
        await provider.close()
        assert second_poller_task.done()
        assert sampler._futures_poller.closed_tasks == [
            first_poller_task,
            second_poller_task,
        ]
        assert holder.cleanup_calls == 2

    try:
        asyncio.run(exercise())
    finally:
        holder.stop()

    assert sampler.close_calls == 0
    assert holder.close_calls == 0
    assert len(services) == 2
    assert len(sampler.calls) == 2


def test_checkpoint_sampler_creation_failure_cleans_owned_service(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import tinker

    monkeypatch.setenv(KEY_ENV_VAR, "secret-value")
    monkeypatch.setattr(provider_module, "version", lambda package: "0.27.1")
    holder = FakeTinkerLoopHolder()

    class FakeServiceClient:
        def __init__(self, **kwargs: object) -> None:
            del kwargs
            self._session_holder = holder

        def create_sampling_client(self, **kwargs: object) -> object:
            del kwargs
            raise RuntimeError("sampler creation failed")

    monkeypatch.setattr(tinker, "ServiceClient", FakeServiceClient)
    provider = TinkerProvider(
        sampler_checkpoint=(
            "tinker://01234567-89ab-cdef-0123-456789abcdef:train:0/sampler_weights/final"
        ),
        tokenizer=FakeTokenizer('{"cells":[]}'),
    )

    try:
        with pytest.raises(RuntimeError, match="sampler creation failed"):
            asyncio.run(provider.complete(system_prompt="system", user_prompt="user"))
    finally:
        holder.stop()

    assert holder.cleanup_calls == 1
    assert provider._sampling_client is None
    assert provider._service_client is None


def test_provider_close_is_noop_before_native_creation_without_poller_or_when_injected(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv(KEY_ENV_VAR, "secret-value")
    tokenizer = FakeTokenizer('{"cells":[]}')
    uncreated = TinkerProvider(tokenizer=tokenizer)

    class TrapInjectedSampler:
        @property
        def holder(self) -> object:
            raise AssertionError("injected sampler internals must not be inspected")

        @property
        def _futures_poller(self) -> object:
            raise AssertionError("injected sampler internals must not be inspected")

    injected = TinkerProvider(
        sampling_client=TrapInjectedSampler(),
        tokenizer=tokenizer,
    )
    no_poller = TinkerProvider(tokenizer=tokenizer)
    no_poller._sampling_client = SimpleNamespace(
        holder=SimpleNamespace(
            run_coroutine_threadsafe=lambda _coroutine: pytest.fail(
                "a missing poller must not schedule shutdown work"
            )
        ),
        _futures_poller=None,
    )

    async def close_all() -> None:
        await uncreated.close()
        await uncreated.close()
        await injected.close()
        await no_poller.close()

    asyncio.run(close_all())


def test_provider_close_closes_cleanup_coroutine_when_holder_rejects_scheduling(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv(KEY_ENV_VAR, "secret-value")
    monkeypatch.setattr(provider_module, "version", lambda package: "0.27.1")

    class RejectingHolder:
        cleanup: object | None = None

        def run_coroutine_threadsafe(self, coroutine: object) -> object:
            self.cleanup = coroutine
            raise RuntimeError("holder loop is closed")

    holder = RejectingHolder()
    sampler = SimpleNamespace(
        holder=holder,
        _futures_poller=SimpleNamespace(_task=object(), close=lambda: None),
    )
    provider = TinkerProvider(tokenizer=FakeTokenizer('{"cells":[]}'))
    provider._sampling_client = sampler

    with pytest.raises(RuntimeError, match="holder loop is closed"):
        asyncio.run(provider.close())

    assert holder.cleanup is not None
    assert holder.cleanup.cr_frame is None


@pytest.mark.parametrize(
    "tokens,stop_reason,rejection",
    [
        ([201, STOP_TOKEN_ID], "length", "rejected_stop_reason"),
        ([201], "stop", "rejected_stop_token"),
        ([STOP_TOKEN_ID, 201, STOP_TOKEN_ID], "stop", "rejected_stop_token"),
    ],
)
def test_provider_rejects_inexact_termination_with_token_evidence(
    monkeypatch: pytest.MonkeyPatch,
    tokens: list[int],
    stop_reason: str,
    rejection: str,
) -> None:
    monkeypatch.setenv(KEY_ENV_VAR, "secret-value")
    provider = TinkerProvider(
        sampling_client=FakeSampler(decoded_tokens=tokens, stop_reason=stop_reason),
        tokenizer=FakeTokenizer('{"cells":[]}'),
    )

    with pytest.raises(ProviderResponseError) as caught:
        asyncio.run(provider.complete(system_prompt="system", user_prompt="user"))

    assert caught.value.stop_reason == stop_reason
    assert caught.value.response_rejection == rejection
    assert caught.value.input_tokens == 3
    assert caught.value.output_tokens == len(tokens)


def test_provider_rejects_multiple_sequences(monkeypatch: pytest.MonkeyPatch) -> None:
    sampler = FakeSampler(decoded_tokens=[201, STOP_TOKEN_ID])
    sampler.response.sequences.append(
        SimpleNamespace(tokens=[202, STOP_TOKEN_ID], stop_reason="stop")
    )
    monkeypatch.setenv(KEY_ENV_VAR, "secret-value")
    provider = TinkerProvider(sampling_client=sampler, tokenizer=FakeTokenizer('{"cells":[]}'))

    with pytest.raises(ProviderResponseError) as caught:
        asyncio.run(provider.complete(system_prompt="system", user_prompt="user"))

    assert caught.value.response_rejection == "rejected_sequence_count"


def test_non_retryable_failure_opens_safe_local_circuit(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    class FakeAuthenticationError(RuntimeError):
        status_code = 401

        def __str__(self) -> str:
            return "secret response body"

    class FailingSampler:
        def __init__(self) -> None:
            self.calls = 0

        async def sample_async(self, **kwargs: object) -> object:
            self.calls += 1
            raise FakeAuthenticationError()

    sampler = FailingSampler()
    monkeypatch.setenv(KEY_ENV_VAR, "secret-value")
    provider = TinkerProvider(sampling_client=sampler, tokenizer=FakeTokenizer("unused"))

    for _ in range(2):
        with pytest.raises(ProviderConfigurationError) as caught:
            asyncio.run(provider.complete(system_prompt="system", user_prompt="user"))
        assert "HTTP 401" in str(caught.value)
        assert "FakeAuthenticationError" in str(caught.value)
        assert "secret" not in str(caught.value)
    assert sampler.calls == 1


def test_transient_provider_failure_is_not_retried(monkeypatch: pytest.MonkeyPatch) -> None:
    class FakeServerError(RuntimeError):
        status_code = 500

    class FailingSampler:
        def __init__(self) -> None:
            self.calls = 0

        async def sample_async(self, **kwargs: object) -> object:
            self.calls += 1
            raise FakeServerError()

    sampler = FailingSampler()
    monkeypatch.setenv(KEY_ENV_VAR, "secret-value")
    provider = TinkerProvider(sampling_client=sampler, tokenizer=FakeTokenizer("unused"))

    with pytest.raises(FakeServerError):
        asyncio.run(provider.complete(system_prompt="system", user_prompt="user"))
    assert sampler.calls == 1
