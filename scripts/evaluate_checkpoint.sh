#!/usr/bin/env bash
set -euo pipefail

repo_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$repo_root"

if [[ $# -lt 2 || $# -gt 4 || "$1" != "--execute" ]]; then
  echo "usage: $0 --execute tinker://SAMPLER_CHECKPOINT [NEW_OUTPUT_ROOT] [DATASET_DIR]" >&2
  exit 2
fi

sampler_checkpoint="$2"
output_root="${3:-training/runs/checkpoint-evaluation-$(date -u +%Y%m%dT%H%M%SZ)}"
dataset_dir="${4:-${FORMULABENCH_DATASET_DIR:-data/spreadsheetbench_verified_400}}"

if [[ "$sampler_checkpoint" != tinker://* ]]; then
  echo "sampler checkpoint must be a tinker:// path" >&2
  exit 2
fi
if [[ -z "${TINKER_API_KEY:-}" ]]; then
  echo "TINKER_API_KEY is required and must be supplied through the environment" >&2
  exit 2
fi
if [[ -z "${TINKER_PROJECT_ID:-}" ]]; then
  echo "TINKER_PROJECT_ID is required and must identify the checkpoint's writable project" >&2
  exit 2
fi
if [[ -e "$output_root" || -L "$output_root" ]]; then
  echo "output root must be a new path" >&2
  exit 2
fi

env -u TINKER_API_KEY uv sync --locked --extra native-tinker

# Apply the production checkpoint-path and runtime validation before corpus
# rebuilding or either paid inference arm. Preflight-only creates no output.
env -u TINKER_API_KEY uv run python -m formulabench.cli \
  --dataset-dir "$dataset_dir" \
  --out-dir "$output_root/checkpoint" \
  --sampler-checkpoint "$sampler_checkpoint" \
  --preflight-only

env -u TINKER_API_KEY uv run python -m training.corpus build \
  --dataset-dir "$dataset_dir" \
  --split-manifest experiments/public_split.json \
  --out-dir training/generated
env -u TINKER_API_KEY uv run python -m training.corpus validate \
  --dataset-dir "$dataset_dir" \
  --split-manifest experiments/public_split.json \
  --corpus training/generated/corpus.jsonl \
  --manifest training/generated/corpus_manifest.json

validation_ids="$(
  env -u TINKER_API_KEY uv run python -m training.partition_ids \
    --manifest training/generated/corpus_manifest.json \
    --split-manifest experiments/public_split.json \
    --partition validation
)"
if [[ -z "$validation_ids" ]]; then
  echo "validated SFT corpus has no validation task IDs" >&2
  exit 2
fi

uv run python -m formulabench.cli \
  --dataset-dir "$dataset_dir" \
  --out-dir "$output_root/base" \
  --ids "$validation_ids"
env -u TINKER_API_KEY uv run python evaluate.py \
  --dataset-dir "$dataset_dir" \
  --predictions "$output_root/base/predictions.jsonl" \
  --ids "$validation_ids" \
  --out "$output_root/base-results.json"

uv run python -m formulabench.cli \
  --dataset-dir "$dataset_dir" \
  --out-dir "$output_root/checkpoint" \
  --ids "$validation_ids" \
  --sampler-checkpoint "$sampler_checkpoint"
env -u TINKER_API_KEY uv run python evaluate.py \
  --dataset-dir "$dataset_dir" \
  --predictions "$output_root/checkpoint/predictions.jsonl" \
  --ids "$validation_ids" \
  --out "$output_root/checkpoint-results.json"

env -u TINKER_API_KEY uv run python -m training.compare_results \
  --base "$output_root/base-results.json" \
  --base-predictions "$output_root/base/predictions.jsonl" \
  --base-traces "$output_root/base/traces" \
  --checkpoint "$output_root/checkpoint-results.json" \
  --checkpoint-predictions "$output_root/checkpoint/predictions.jsonl" \
  --checkpoint-traces "$output_root/checkpoint/traces" \
  --sampler-checkpoint "$sampler_checkpoint" \
  --manifest training/generated/corpus_manifest.json \
  --split-manifest experiments/public_split.json \
  --out "$output_root/comparison.json"

echo "comparison written to $output_root/comparison.json"
