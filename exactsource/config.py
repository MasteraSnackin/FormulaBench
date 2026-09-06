"""Reproducible inference settings.

The hackathon contract requires model identifiers and inference settings to be
fixed in source rather than selected by the judge at runtime.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

MODEL_ID = "Qwen/Qwen3.8-27B"
MODEL_NAME = f"tinker:{MODEL_ID}"
API_KEY_ENV = "TINKER_API_KEY"
TINKER_ANTHROPIC_BASE_URL = "https://tinker.thinkingmachines.dev/services/tinker-prod/anthropic/api"
TINKER_MESSAGES_URL = f"{TINKER_ANTHROPIC_BASE_URL}/v1/messages"
TEMPERATURE = 0.0
# Cell plans are compact declarative JSON, so a 16k cap bounds cost while leaving
# enough room for Qwen's reasoning prelude and the complete operations payload.
CELL_MAX_OUTPUT_TOKENS = 16_000
# A capped cell response gets one no-think-requested recovery with more room to
# finish the operations payload; this does not change the ordinary cell budget.
CELL_TRUNCATION_RECOVERY_MAX_OUTPUT_TOKENS = 32_000
# Sheet transformations can contain substantial Python source and therefore use a
# larger cap on both the initial generation and an ordinary semantic repair.
SHEET_MAX_OUTPUT_TOKENS = 32_000
# Preserve the public client default for callers that do not select a task route.
MAX_OUTPUT_TOKENS = CELL_MAX_OUTPUT_TOKENS
REASONING_EFFORT = True
# The cell recovery also adds Qwen's /no_think switch and an empty-thinking
# prefill to bias the larger allowance towards the answer.
CELL_TRUNCATION_RECOVERY_REASONING_EFFORT = False
MODEL_CONNECT_TIMEOUT_SECONDS = 20.0
MODEL_STREAM_READ_TIMEOUT_SECONDS = 300.0
MODEL_WRITE_TIMEOUT_SECONDS = 30.0
MODEL_POOL_TIMEOUT_SECONDS = 30.0
TRANSPORT_RETRIES = 2
MAX_RETRY_AFTER_SECONDS = 60.0
MAX_RETRY_STAGGER_SECONDS = 0.25
SEMANTIC_REPAIRS = 1
CONCURRENCY = 4
CONTEXT_CHAR_BUDGET = 48_000
TRACE_TEXT_LIMIT = 20_000


@dataclass(frozen=True, slots=True)
class RuntimePaths:
    """Judge mount points and their derived output paths."""

    data_dir: Path = Path("/data")
    out_dir: Path = Path("/out")

    @property
    def outputs_dir(self) -> Path:
        return self.out_dir / "outputs"

    @property
    def traces_dir(self) -> Path:
        return self.out_dir / "traces"

    @property
    def predictions_path(self) -> Path:
        return self.out_dir / "predictions.jsonl"

    @property
    def log_path(self) -> Path:
        return self.out_dir / "run.log"
