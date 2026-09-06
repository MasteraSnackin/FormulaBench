#!/bin/sh

set -eu

dataset_dir=${1:-data/spreadsheetbench_verified_400}
reference_run=${2:-../ExactSource}
image_name=${FORMULABENCH_PARITY_IMAGE:-formulabench:contract-test}

if ! command -v docker >/dev/null 2>&1; then
    echo "Docker is required but was not found." >&2
    exit 2
fi
if [ ! -f "$dataset_dir/dataset.json" ]; then
    echo "Dataset not found at $dataset_dir. Run: uv run python data/download.py" >&2
    exit 2
fi
if [ ! -f "$reference_run/predictions.jsonl" ]; then
    echo "ExactSource reference run not found at $reference_run." >&2
    exit 2
fi

dataset_path=$(cd "$dataset_dir" && pwd -P)
reference_path=$(cd "$reference_run" && pwd -P)

docker build --target contract-test --tag "$image_name" .
docker run --rm \
    --network none \
    --read-only \
    --user 12345:12345 \
    --tmpfs /tmp:rw,nosuid,nodev,size=1g \
    --env PYTHONDONTWRITEBYTECODE=1 \
    --mount "type=bind,src=$dataset_path,dst=/dataset,readonly" \
    --mount "type=bind,src=$reference_path,dst=/reference,readonly" \
    --entrypoint /app/.venv/bin/python \
    "$image_name" \
    tools/verify_v2_parity.py \
    --dataset-dir /dataset \
    --reference-run /reference
