# syntax=docker/dockerfile:1.7
ARG PYTHON_IMAGE=python:3.11-slim-bookworm@sha256:528257d48c1da0dcecc2e725d1ae34498d60c965f1241e39cd6a85a8859bdf84
ARG UV_IMAGE=ghcr.io/astral-sh/uv:0.12.10@sha256:2bb3ebca0a796a155094a27773d290c4b074572e6107f171d88d086682fd2500

FROM ${UV_IMAGE} AS uv-bin

FROM ${PYTHON_IMAGE} AS base

FROM base AS build

COPY --from=uv-bin /uv /usr/local/bin/uv

ENV UV_LINK_MODE=copy \
    UV_PYTHON_DOWNLOADS=never \
    HF_HOME=/app/.cache/huggingface

WORKDIR /app

COPY pyproject.toml uv.lock ./
RUN uv sync --locked --no-dev --no-install-project

RUN .venv/bin/python - <<'PY'
import hashlib
from pathlib import Path

from huggingface_hub import snapshot_download
from transformers import AutoTokenizer

revision = "1d4bf0f2ff6012fd82039f2fa52739d0dd7c60c0"
expected = {
    "chat_template.jinja": "c3cf9e34abf4f9e36c2d72165aa9c132d3e2a725b6c2586aaa3a8af9d7a81041",
    "config.json": "191e0af232104ed8b65258cf3fb2b842e288008baca7633c11b82a1ac7203aab",
    "tokenizer.json": "0997f410c57a1f4e53b09e4be8f4a172d90edd9564368fb0847030937229b9f3",
    "tokenizer_config.json": "b11349aafa7cdc6a320767cf7ceb29ed82f7eda5d65e8e0819e76f0ce947bf27",
}
snapshot = Path(
    snapshot_download(
        repo_id="Qwen/Qwen3.8-27B",
        revision=revision,
        allow_patterns=list(expected),
    )
)
for name, digest in expected.items():
    actual = hashlib.sha256((snapshot / name).read_bytes()).hexdigest()
    if actual != digest:
        raise RuntimeError(f"pinned tokenizer asset failed verification: {name}")
tokenizer = AutoTokenizer.from_pretrained(
    "Qwen/Qwen3.8-27B",
    revision=revision,
    local_files_only=True,
)
if tokenizer.convert_tokens_to_ids("<|im_end|>") != 248046:
    raise RuntimeError("pinned tokenizer stop token failed verification")
PY

FROM build AS contract-test

COPY sb.py ./
COPY formulabench/ ./formulabench/
COPY experiments/ ./experiments/
COPY tests/ ./tests/
RUN uv sync --locked --no-install-project
RUN .venv/bin/python -m pytest

FROM base AS runtime

RUN groupadd --gid 10001 runner \
    && useradd --uid 10001 --gid 10001 --create-home --shell /usr/sbin/nologin runner \
    && install -d -m 0555 /data \
    && install -d -m 0755 -o 10001 -g 10001 /out

WORKDIR /app

ENV VIRTUAL_ENV=/app/.venv \
    PATH=/app/.venv/bin:/usr/local/bin:/usr/bin:/bin \
    PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    XDG_CACHE_HOME=/tmp/.cache \
    HF_HOME=/app/.cache/huggingface \
    HF_HUB_DISABLE_TELEMETRY=1 \
    HF_HUB_OFFLINE=1 \
    TINKER_TELEMETRY=0 \
    TRANSFORMERS_OFFLINE=1 \
    TRANSFORMERS_VERBOSITY=error

COPY --from=build --chown=10001:10001 /app/.venv/ /app/.venv/
COPY --from=build --chown=10001:10001 /app/.cache/ /app/.cache/
# The wrapper deliberately runs with the host UID so bind-mounted results are
# host-owned.  Tokenizer assets must therefore be readable by an arbitrary,
# unprivileged runtime UID rather than only by the image's default user.
RUN chmod -R a+rX /app/.cache
COPY --chown=10001:10001 sb.py ./
COPY --chown=10001:10001 formulabench/ ./formulabench/
COPY --chown=10001:10001 scripts/container_entrypoint.sh ./container_entrypoint.sh

USER 10001:10001

ENTRYPOINT ["/app/container_entrypoint.sh"]
