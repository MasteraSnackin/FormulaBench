#!/bin/sh

set -eu

image_name=${FORMULABENCH_IMAGE:-formulabench:latest}
dataset_dir=data/spreadsheetbench_verified_400
output_dir=submissions/formulabench-v2

normalise_output_leaf() {
    candidate=$1
    while [ "$candidate" != "/" ]; do
        case "$candidate" in
            */) candidate=${candidate%/} ;;
            */.)
                candidate=${candidate%/.}
                if [ -z "$candidate" ]; then
                    candidate=/
                fi
                ;;
            *) break ;;
        esac
    done
    printf '%s\n' "$candidate"
}

resolve_prospective_directory() {
    candidate=$1
    case "$candidate" in
        /*) absolute=$candidate ;;
        *) absolute=$(pwd -P)/$candidate ;;
    esac

    probe=$absolute
    while [ ! -d "$probe" ]; do
        parent=$(dirname "$probe")
        if [ "$parent" = "$probe" ]; then
            return 1
        fi
        probe=$parent
    done

    resolved=$(cd "$probe" && pwd -P) || return 1
    remaining=${absolute#"$probe"}
    remaining=${remaining#/}
    while [ -n "$remaining" ]; do
        component=${remaining%%/*}
        if [ "$component" = "$remaining" ]; then
            remaining=
        else
            remaining=${remaining#*/}
        fi
        case "$component" in
            ""|.) ;;
            ..)
                if [ "$resolved" != "/" ]; then
                    resolved=${resolved%/*}
                    if [ -z "$resolved" ]; then
                        resolved=/
                    fi
                fi
                ;;
            *)
                if [ "$resolved" = "/" ]; then
                    resolved=/$component
                else
                    resolved=$resolved/$component
                fi
                ;;
        esac
    done
    printf '%s\n' "$resolved"
}

path_is_within() (
    child=$1
    parent=$2
    if [ "$child" = "$parent" ] || [ "$parent" = "/" ]; then
        return 0
    fi
    case "$child" in
        "$parent"/*) return 0 ;;
        *) return 1 ;;
    esac
)

for argument in "$@"; do
    case "$argument" in
        --dataset-dir|--dataset-dir=*|--out-dir|--out-dir=*)
            echo "--dataset-dir and --out-dir are managed by this wrapper; pass dataset and output directories as the first two positional arguments." >&2
            exit 2
            ;;
    esac
done

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
legacy_engine=0
legacy_only_option=0
for argument in "$@"; do
    case "$argument" in
        --legacy-engine) legacy_engine=1 ;;
        --resume) resume=1; legacy_only_option=1 ;;
        --retry-failures) legacy_only_option=1 ;;
        --preflight-only) preflight=1 ;;
        --replay-write-failures) replay_write_failures=1; legacy_only_option=1 ;;
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
dataset_path=$(cd "$dataset_dir" && pwd -P)

if [ "$replay_write_failures" -eq 1 ] && [ "$resume" -eq 0 ]; then
    echo "--replay-write-failures requires --resume." >&2
    exit 2
fi

if [ "$legacy_only_option" -eq 1 ] && [ "$legacy_engine" -eq 0 ]; then
    echo "Resume, retry and replay options require --legacy-engine; v2 supports fresh runs only." >&2
    exit 2
fi

if [ "$preflight" -eq 0 ] && [ "$replay_write_failures" -eq 0 ]; then
    case "${TINKER_API_KEY:-}" in
        *[![:space:]]*) ;;
        *)
            echo "TINKER_API_KEY is not set. Supply it privately through your shell environment." >&2
            exit 2
            ;;
    esac
fi

if [ -z "$output_dir" ]; then
    echo "Output directory must not be empty." >&2
    exit 2
fi
output_leaf=$(normalise_output_leaf "$output_dir")
if [ -L "$output_leaf" ]; then
    echo "Output directory must not be a symbolic link: $output_dir" >&2
    exit 2
fi
if ! output_path=$(resolve_prospective_directory "$output_leaf"); then
    echo "Output directory cannot be resolved: $output_dir" >&2
    exit 2
fi
if path_is_within "$output_path" "$dataset_path" || path_is_within "$dataset_path" "$output_path"; then
    echo "Dataset and output directories must not overlap." >&2
    exit 2
fi

if [ "$resume" -eq 1 ]; then
    if [ ! -d "$output_path" ] || [ ! -f "$output_path/run.log" ]; then
        echo "Resume requires an existing FormulaBench output directory: $output_dir" >&2
        exit 2
    fi
else
    if [ -e "$output_path" ] && [ ! -d "$output_path" ]; then
        echo "Output path is not a directory: $output_dir" >&2
        exit 2
    fi
    mkdir -p "$output_path"
    if [ -n "$(find "$output_path" -mindepth 1 -maxdepth 1 -print -quit)" ]; then
        echo "A fresh run requires an empty output directory: $output_dir" >&2
        exit 2
    fi
fi

docker build --target runtime --tag "$image_name" .
if [ "$preflight" -eq 1 ] || [ "$replay_write_failures" -eq 1 ]; then
    docker run --rm \
        --user "$(id -u):$(id -g)" \
        --mount "type=bind,src=$dataset_path,dst=/data,readonly" \
        --mount "type=bind,src=$output_path,dst=/out" \
        "$image_name" \
        "$@"
elif [ "$legacy_engine" -eq 1 ]; then
    docker run --rm \
        --user "$(id -u):$(id -g)" \
        --env TINKER_API_KEY \
        --env TINKER_PROJECT_ID \
        --mount "type=bind,src=$dataset_path,dst=/data,readonly" \
        --mount "type=bind,src=$output_path,dst=/out" \
        "$image_name" \
        "$@"
else
    docker run --rm \
        --user "$(id -u):$(id -g)" \
        --env TINKER_API_KEY \
        --mount "type=bind,src=$dataset_path,dst=/data,readonly" \
        --mount "type=bind,src=$output_path,dst=/out" \
        "$image_name" \
        "$@"
fi

echo "FormulaBench output: $output_path"
