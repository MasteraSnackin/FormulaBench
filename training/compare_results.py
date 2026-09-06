"""Create a hash-bound base-versus-checkpoint evaluator comparison."""

from __future__ import annotations

import argparse
import hashlib
import json
import math
import os
import tempfile
from collections.abc import Mapping, Sequence
from pathlib import Path
from typing import Any

from formulabench.constants import MODEL_ID
from formulabench.provider import ProviderConfigurationError, validate_sampler_checkpoint
from training.train_lora import (
    DEFAULT_CORPUS_MANIFEST,
    DEFAULT_SPLIT_MANIFEST,
    TrainingInputError,
    load_validated_corpus,
)

MAX_RESULTS_BYTES = 32 * 1024 * 1024
METRICS = ("pass_rate", "cell_accuracy", "pass_rate_cell_level", "pass_rate_sheet_level")


class ComparisonError(ValueError):
    """An evaluator result cannot support a controlled comparison."""


def _load_result(path: Path, label: str) -> tuple[Mapping[str, Any], bytes]:
    if path.is_symlink() or not path.is_file():
        raise ComparisonError(f"{label} result must be a regular non-symlink file")
    raw = path.read_bytes()
    if not raw or len(raw) > MAX_RESULTS_BYTES:
        raise ComparisonError(f"{label} result has an invalid size")
    try:
        payload = json.loads(raw.decode("utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise ComparisonError(f"{label} result is not valid UTF-8 JSON") from exc
    if not isinstance(payload, Mapping):
        raise ComparisonError(f"{label} result root must be an object")
    summary = payload.get("summary")
    items = payload.get("items")
    if not isinstance(summary, Mapping) or not isinstance(items, list) or not items:
        raise ComparisonError(f"{label} result lacks a summary or item records")
    for metric in METRICS:
        value = summary.get(metric)
        if value is not None and (
            isinstance(value, bool)
            or not isinstance(value, (int, float))
            or not math.isfinite(value)
        ):
            raise ComparisonError(f"{label} result has an invalid {metric}")
    return payload, raw


def _task_ids(payload: Mapping[str, Any], label: str) -> tuple[str, ...]:
    ids: list[str] = []
    for index, raw_item in enumerate(payload["items"]):
        if not isinstance(raw_item, Mapping):
            raise ComparisonError(f"{label} item {index} is not an object")
        task_id = raw_item.get("id")
        if not isinstance(task_id, str) or not task_id:
            raise ComparisonError(f"{label} item {index} has an invalid task ID")
        ids.append(task_id)
    if len(ids) != len(set(ids)):
        raise ComparisonError(f"{label} result contains duplicate task IDs")
    return tuple(ids)


def compare_results(
    base_path: Path,
    checkpoint_path: Path,
    *,
    expected_task_ids: Sequence[str] | None = None,
) -> dict[str, Any]:
    base, base_raw = _load_result(base_path, "base")
    checkpoint, checkpoint_raw = _load_result(checkpoint_path, "checkpoint")
    base_ids = _task_ids(base, "base")
    checkpoint_ids = _task_ids(checkpoint, "checkpoint")
    if base_ids != checkpoint_ids:
        raise ComparisonError("base and checkpoint results do not contain the same ordered tasks")
    if expected_task_ids is not None and base_ids != tuple(expected_task_ids):
        raise ComparisonError("comparison tasks do not match the validated corpus partition")

    base_summary = dict(base["summary"])
    checkpoint_summary = dict(checkpoint["summary"])
    deltas: dict[str, float | None] = {}
    for metric in METRICS:
        base_value = base_summary.get(metric)
        checkpoint_value = checkpoint_summary.get(metric)
        deltas[metric] = (
            round(float(checkpoint_value) - float(base_value), 4)
            if base_value is not None and checkpoint_value is not None
            else None
        )
    return {
        "schema_version": 1,
        "task_count": len(base_ids),
        "ordered_task_ids": list(base_ids),
        "base_results_sha256": hashlib.sha256(base_raw).hexdigest(),
        "checkpoint_results_sha256": hashlib.sha256(checkpoint_raw).hexdigest(),
        "base": base_summary,
        "checkpoint": checkpoint_summary,
        "checkpoint_minus_base": deltas,
    }


def _write_new_json(path: Path, payload: Mapping[str, Any]) -> None:
    if path.exists() or path.is_symlink():
        raise ComparisonError("comparison output must be a new path")
    path.parent.mkdir(parents=True, exist_ok=True)
    descriptor, temporary_name = tempfile.mkstemp(
        dir=path.parent,
        prefix=f".{path.name}.",
        suffix=".tmp",
    )
    temporary = Path(temporary_name)
    try:
        with os.fdopen(descriptor, "w", encoding="utf-8") as handle:
            json.dump(payload, handle, ensure_ascii=False, indent=2, sort_keys=True)
            handle.write("\n")
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(temporary, path)
    finally:
        temporary.unlink(missing_ok=True)


def parse_args(argv: Sequence[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--base", required=True, type=Path)
    parser.add_argument("--checkpoint", required=True, type=Path)
    parser.add_argument("--sampler-checkpoint", required=True)
    parser.add_argument("--manifest", type=Path, default=DEFAULT_CORPUS_MANIFEST)
    parser.add_argument("--split-manifest", type=Path, default=DEFAULT_SPLIT_MANIFEST)
    parser.add_argument("--out", required=True, type=Path)
    return parser.parse_args(argv)


def main(argv: Sequence[str] | None = None) -> int:
    args = parse_args(argv)
    try:
        try:
            validate_sampler_checkpoint(args.sampler_checkpoint)
        except ProviderConfigurationError as exc:
            raise ComparisonError(str(exc)) from None
        corpus = load_validated_corpus(args.manifest, args.split_manifest)
        validation_ids = [row.task_id for row in corpus.validation_rows]
        comparison = compare_results(
            args.base,
            args.checkpoint,
            expected_task_ids=validation_ids,
        )
        comparison["experiment"] = {
            "base_model": MODEL_ID,
            "corpus_sha256": corpus.corpus_sha256,
            "dataset_manifest_sha256": corpus.dataset_manifest_sha256,
            "split_manifest_sha256": corpus.split_manifest_sha256,
            "sampler_checkpoint_sha256": hashlib.sha256(
                args.sampler_checkpoint.encode("utf-8")
            ).hexdigest(),
        }
        _write_new_json(args.out, comparison)
    except (ComparisonError, OSError, TrainingInputError) as exc:
        raise SystemExit(f"comparison error: {exc}") from exc
    print(json.dumps(comparison["checkpoint_minus_base"], sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
