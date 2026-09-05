#!/bin/sh

set -eu

image_name=${FORMULABENCH_IMAGE:-formulabench:latest}
dataset_dir=data/spreadsheetbench_verified_400
output_dir=submissions/formulabench

if [ "$#" -gt 0 ]; then
    dataset_dir=$1
    shift
fi
if [ "$#" -gt 0 ]; then
    output_dir=$1
    shift
fi

resume=0
preflight=0
replay_write_failures=0
for argument in "$@"; do
    case "$argument" in
        --resume) resume=1 ;;
        --preflight-only) preflight=1 ;;
        --replay-write-failures) replay_write_failures=1 ;;
    esac
done

if ! command -v docker >/dev/null 2>&1; then
    echo "Docker is required but was not found." >&2
    exit 2
fi

if [ ! -f "$dataset_dir/dataset.json" ]; then
    echo "Dataset not found at $dataset_dir. Run: uv run python data/download.py" >&2
    exit 2
fi

if [ "$replay_write_failures" -eq 1 ] && [ "$resume" -eq 0 ]; then
    echo "--replay-write-failures requires --resume." >&2
    exit 2
fi

if [ "$preflight" -eq 0 ] && [ "$replay_write_failures" -eq 0 ] && [ -z "${TINKER_API_KEY:-}" ]; then
    echo "TINKER_API_KEY is not set. Supply it privately through your shell environment." >&2
    exit 2
fi

if [ -L "$output_dir" ]; then
    echo "Output directory must not be a symbolic link: $output_dir" >&2
    exit 2
fi
if [ "$resume" -eq 1 ]; then
    if [ ! -d "$output_dir" ] || [ ! -f "$output_dir/run.log" ]; then
        echo "Resume requires an existing FormulaBench output directory: $output_dir" >&2
        exit 2
    fi
else
    mkdir -p "$output_dir"
    if [ -n "$(find "$output_dir" -mindepth 1 -maxdepth 1 -print -quit)" ]; then
        echo "A fresh run requires an empty output directory: $output_dir" >&2
        exit 2
    fi
fi

dataset_path=$(cd "$dataset_dir" && pwd -P)
output_path=$(cd "$output_dir" && pwd -P)

docker build --target runtime --tag "$image_name" .
if [ "$replay_write_failures" -eq 1 ]; then
    docker run --rm \
        --user "$(id -u):$(id -g)" \
        --mount "type=bind,src=$dataset_path,dst=/data,readonly" \
        --mount "type=bind,src=$output_path,dst=/out" \
        "$image_name" \
        "$@"
else
    docker run --rm \
        --user "$(id -u):$(id -g)" \
        --env TINKER_API_KEY \
        --env TINKER_PROJECT_ID \
        --mount "type=bind,src=$dataset_path,dst=/data,readonly" \
        --mount "type=bind,src=$output_path,dst=/out" \
        "$image_name" \
        "$@"
fi

echo "FormulaBench output: $output_path"
