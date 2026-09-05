# Submission: <team name>

Copy this file to the root of your repo as `SUBMISSION.md` and fill every section. By Sunday 12:00 you send us the link to your repo through the form we share.

## Team

- Team name:
- Members, one GitHub handle per line:
- Repo URL:

## What we built and why

One paragraph, 150 to 300 words. The problem you picked, what you built, why you chose that route, what worked, what did not. Be exact: which model, which prompt strategy, which training data, which tools the agent had. This paragraph is what the judges read first.

## Models

Every model your pipeline calls at inference time: the exact model id, or for a fine-tune the `tinker://` sampler checkpoint path plus the base model, with the TTL cleared. If you fine-tuned: training data and how you built it, whether the 400 golden files were in it, steps, learning rate, compute, wall time.

## Scores on the 400

Produced by the shipped evaluator, nothing else:

```sh
uv run evaluate.py --predictions <your predictions.jsonl> --all --out results.json
```

Paste the `summary` block of `results.json` here and put the file in the repo. `items` must be 400.

```json
{"items": 400, "graded": ..., "missing": ..., "errors": ..., "pass_rate": ..., "cell_accuracy": ..., "pass_rate_cell_level": ..., "pass_rate_sheet_level": ...}
```

## Your run on the 400

In the repo:

- `predictions.jsonl`: path
- `outputs/`: path. Whatever your pipeline is, one model call or an agent, the deliverable is one workbook per task. The evaluator reads `predictions.jsonl` to find each one and grades only the answer cells in it.
- `traces/`: path, one `<id>.jsonl` per task, one line per model call
- `run.log`: path

## Code

The code that produced the run is in the repo. If your pipeline is an agent or executes model-written code, the repo has the Dockerfile it ran in, reading `/data` and writing `/out`. Name the environment variables it needs.

## Things to look at

Notebooks, experiment logs, ablations, failure analyses, training curves, anything that shows how you got here. One line each: path, and what we will find there.

-
-
