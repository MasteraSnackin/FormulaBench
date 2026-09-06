"""Print a validated FormulaBench SFT partition as comma-separated task IDs."""

from __future__ import annotations

import argparse
import sys
from collections.abc import Sequence
from pathlib import Path

from training.train_lora import (
    DEFAULT_CORPUS_MANIFEST,
    DEFAULT_SPLIT_MANIFEST,
    TrainingInputError,
    load_validated_corpus,
)


def parse_args(argv: Sequence[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--manifest", type=Path, default=DEFAULT_CORPUS_MANIFEST)
    parser.add_argument("--split-manifest", type=Path, default=DEFAULT_SPLIT_MANIFEST)
    parser.add_argument("--partition", choices=("train", "validation"), required=True)
    return parser.parse_args(argv)


def main(argv: Sequence[str] | None = None) -> int:
    args = parse_args(argv)
    try:
        corpus = load_validated_corpus(args.manifest, args.split_manifest)
    except TrainingInputError as exc:
        print(f"partition preflight failed: {exc}", file=sys.stderr)
        return 2
    rows = corpus.train_rows if args.partition == "train" else corpus.validation_rows
    print(",".join(row.task_id for row in rows))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
