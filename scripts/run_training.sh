#!/usr/bin/env bash
set -euo pipefail

repo_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$repo_root"
export TINKER_TELEMETRY=0

dataset_dir="${FORMULABENCH_DATASET_DIR:-data/spreadsheetbench_verified_400}"
if [[ $# -gt 0 && "$1" != --* ]]; then
  dataset_dir="$1"
  shift
fi

execute=0
for argument in "$@"; do
  case "$argument" in
    --execute)
      execute=1
      ;;
    --manifest|--manifest=*|--split-manifest|--split-manifest=*)
      echo "run_training.sh fixes the corpus and split manifests; overrides are not allowed" >&2
      exit 2
      ;;
  esac
done

env -u TINKER_API_KEY uv sync --locked --extra native-tinker
env -u TINKER_API_KEY uv run python -m training.corpus build \
  --dataset-dir "$dataset_dir" \
  --split-manifest experiments/public_split.json \
  --out-dir training/generated
env -u TINKER_API_KEY uv run python -m training.corpus validate \
  --dataset-dir "$dataset_dir" \
  --split-manifest experiments/public_split.json \
  --corpus training/generated/corpus.jsonl \
  --manifest training/generated/corpus_manifest.json
if [[ "$execute" -eq 1 ]]; then
  uv run --extra native-tinker python -m training.train_lora \
    "$@" \
    --manifest training/generated/corpus_manifest.json \
    --split-manifest experiments/public_split.json
else
  env -u TINKER_API_KEY uv run --extra native-tinker python -m training.train_lora \
    "$@" \
    --manifest training/generated/corpus_manifest.json \
    --split-manifest experiments/public_split.json
fi
