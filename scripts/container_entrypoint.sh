#!/bin/sh

set -eu

exec /app/.venv/bin/python -m formulabench.v2 \
    "$@" \
    --dataset-dir=/data \
    --out-dir=/out
