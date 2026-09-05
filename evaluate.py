import argparse
import json
import tempfile
from dataclasses import dataclass
from pathlib import Path

from sb import DEFAULT_DATASET, load_answer_values, load_dataset, read_jsonl, recalculate, resolve_output, values_equal


@dataclass(frozen=True)
class EvaluateConfig:
    predictions_path: Path | None
    dataset_dir: Path
    ids: set[str] | None
    recalc: bool
    all_tasks: bool
    oracle: bool
    quiet: bool
    results_path: Path | None


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser()
    p.add_argument("--predictions", help="predictions.jsonl from a baseline run")
    p.add_argument("--dataset-dir", default=str(DEFAULT_DATASET))
    p.add_argument("--ids", help="comma-separated task ids to score")
    p.add_argument("--no-recalc", action="store_true", help="skip LibreOffice recalculation of output files")
    p.add_argument("--all", action="store_true", help="score every task in the dataset dir; a task without a prediction counts as a fail. Judges use this.")
    p.add_argument("--oracle", action="store_true", help="score golden against golden to check the grader")
    p.add_argument("--out", help="write summary and per-item results as JSON")
    p.add_argument("--quiet", action="store_true")
    args = p.parse_args()
    if not args.predictions and not args.oracle:
        p.error("--predictions is required unless --oracle")
    return args


def config_from_args(args: argparse.Namespace) -> EvaluateConfig:
    ids = None
    if args.ids:
        ids = {task_id.strip() for task_id in args.ids.split(",") if task_id.strip()}
    return EvaluateConfig(
        predictions_path=Path(args.predictions) if args.predictions else None,
        dataset_dir=Path(args.dataset_dir),
        ids=ids,
        recalc=not args.no_recalc,
        all_tasks=args.all,
        oracle=args.oracle,
        quiet=args.quiet,
        results_path=Path(args.out) if args.out else None,
    )


def selected_tasks(dataset_dir: Path, ids: set[str] | None) -> list[dict]:
    """Return dataset tasks, optionally filtered to --ids."""
    tasks = load_dataset(dataset_dir)
    if ids is None:
        return tasks
    return [task for task in tasks if task["id"] in ids]


def resolved_ids(config: EvaluateConfig, predictions: list[dict]) -> set[str] | None:
    """Use explicit --ids, otherwise score only tasks present in predictions.jsonl."""
    if config.ids is not None:
        return config.ids
    if config.all_tasks:
        return None
    if predictions and not config.oracle:
        return {prediction["id"] for prediction in predictions}
    return None


def predictions_by_id(predictions: list[dict]) -> dict[str, dict]:
    return {prediction["id"]: prediction for prediction in predictions}


def task_output_path(
    task: dict,
    prediction: dict | None,
    *,
    oracle: bool,
    predictions_path: Path | None,
) -> str | None:
    if oracle:
        return task["golden_xlsx"]
    if prediction is None:
        return None
    return resolve_output(prediction.get("output"), str(predictions_path or "."))


def golden_cell_count(task) -> int:
    """Graded cells of a task, so a missing prediction counts as zero correct in cell_accuracy."""
    if not task["golden_xlsx"]:
        return 0
    try:
        return len(load_answer_values(task["golden_xlsx"], task))
    except Exception:
        return 0


def score_task(task, output_xlsx, recalc, work_dir):
    if not task["golden_xlsx"]:
        return {"status": "no_golden"}
    if not output_xlsx or not Path(output_xlsx).exists():
        return {"status": "missing_output", "cells": golden_cell_count(task), "correct": 0}
    try:
        path = recalculate(output_xlsx, work_dir) if recalc else output_xlsx
        gold = load_answer_values(task["golden_xlsx"], task)
        pred = load_answer_values(path, task)
    except Exception as e:
        return {"status": "error", "error": str(e)[:200]}
    mismatches = [{"cell": f"{k[0]}!{k[1]}", "expected": g, "actual": pred.get(k)}
                  for k, g in gold.items() if not values_equal(g, pred.get(k))]
    return {
        "status": "graded",
        "cells": len(gold),
        "correct": len(gold) - len(mismatches),
        "pass": not mismatches,
        "mismatches": mismatches[:5],
    }


def score(predictions, tasks, *, recalc=True, oracle=False, predictions_path=None):
    by_id = predictions_by_id(predictions)
    items = []
    predictions_path = Path(predictions_path) if predictions_path else None
    with tempfile.TemporaryDirectory() as work:
        for task in tasks:
            prediction = by_id.get(task["id"])
            output = task_output_path(task, prediction, oracle=oracle, predictions_path=predictions_path)
            if not task["golden_xlsx"]:
                items.append({"id": task["id"], "type": task["instruction_type"], "status": "no_golden"})
                continue
            if output is None:
                items.append({"id": task["id"], "type": task["instruction_type"], "status": "missing",
                              "cells": golden_cell_count(task), "correct": 0})
                continue
            result = score_task(task, output, recalc and not oracle, work)
            result.update({"id": task["id"], "type": task["instruction_type"]})
            items.append(result)
    return summarise(items), items


def summarise(items):
    graded = [i for i in items if i["status"] == "graded"]
    counted = [i for i in items if "cells" in i]  # graded plus missing, so a missing task lowers cell_accuracy
    n = len(items)
    if items and all(i["status"] == "no_golden" for i in items):
        return {"items": n, "graded": 0, "no_golden": n, "pass_rate": None, "cell_accuracy": None,
                "note": "no golden files in this dataset dir, nothing was scored"}

    def rate(rows):
        return round(sum(i.get("pass", False) for i in rows) / len(rows), 4) if rows else None

    return {
        "items": n,
        "graded": len(graded),
        "missing": sum(i["status"] in ("missing", "missing_output") for i in items),
        "errors": sum(i["status"] == "error" for i in items),
        "pass_rate": round(sum(i.get("pass", False) for i in items) / n, 4) if n else None,
        "cell_accuracy": round(sum(i["correct"] for i in counted) / sum(i["cells"] for i in counted), 4) if sum(i["cells"] for i in counted) else None,
        "pass_rate_cell_level": rate([i for i in items if i["type"].startswith("Cell")]),
        "pass_rate_sheet_level": rate([i for i in items if i["type"].startswith("Sheet")]),
    }


def print_report(items, summary) -> None:
    print(f"{'id':<8} {'type':<6} {'ok':<4} {'cells':>9}  first mismatch")
    for item in items:
        if item["status"] != "graded":
            print(f"{item['id']:<8} {item['type'][:5]:<6} {'--':<4} {'':>9}  {item['status']} {item.get('error', '')}")
            continue
        mark = "PASS" if item["pass"] else "."
        first = item["mismatches"][0] if item["mismatches"] else None
        detail = f"{first['cell']}: expected {first['expected']!r}, got {first['actual']!r}" if first else ""
        print(f"{item['id']:<8} {item['type'][:5]:<6} {mark:<4} {item['correct']:>4}/{item['cells']:<4}  {detail[:70]}")
    print()
    print(json.dumps(summary, indent=2))


def write_results(path: Path, summary, items) -> None:
    path.write_text(json.dumps({"summary": summary, "items": items}, indent=2, default=str))


def main():
    config = config_from_args(parse_args())
    predictions = read_jsonl(config.predictions_path) if config.predictions_path else []
    ids = resolved_ids(config, predictions)
    tasks = selected_tasks(config.dataset_dir, ids)
    summary, items = score(
        predictions,
        tasks,
        recalc=config.recalc,
        oracle=config.oracle,
        predictions_path=config.predictions_path,
    )

    if not config.quiet:
        print_report(items, summary)
    if config.results_path:
        write_results(config.results_path, summary, items)


if __name__ == "__main__":
    main()
