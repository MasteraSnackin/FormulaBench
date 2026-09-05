#!/bin/sh

set -eu

exec /app/.venv/bin/python -m formulabench.capture \
    --dataset-dir=/data \
    --out-dir=/out \
    "$@"
