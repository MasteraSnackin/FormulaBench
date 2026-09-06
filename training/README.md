# Tinker training experiment

This directory contains a reproducible, development-only LoRA supervised fine-tuning experiment
for `Qwen/Qwen3.8-27B`. It is an optional research path. It does not change, replace or retroactively
improve the frozen 400-task FormulaBench result in `submissions/formulabench/`.

## What the guardrails enforce

- The checked-in public split must reproduce exactly from `dataset.json` and all 400 initial
  workbooks before any golden workbook is located.
- Only the 80 development IDs may supply golden answers. The 80 held-out and 240 final-only golden
  workbooks are not opened by the corpus builder.
- Six development tasks with more than 500 target cells are recorded as explicit exclusions. No
  target or label is silently truncated.
- The remaining 74 tasks are deterministically stratified into 59 training and 15 validation
  examples. A complete task belongs to one partition only.
- Every initial workbook, development golden, prompt, answer, row, corpus and source manifest is
  bound by SHA-256 in the generated manifest.
- Every answer must pass FormulaBench's response, target-coverage, formula-safety and workbook-write
  contracts before it becomes a label.
- Qwen tool declarations and tool-call labels use the production `submit_spreadsheet_answer`
  schema. Preflight requires token-for-token prompt parity with production inference.

Generated corpora and local run records are deliberately ignored by Git. They can contain public
golden answer content and provider checkpoint paths, so they should be regenerated from their
verified sources rather than committed.

## Safe local preflight

The convenience script installs the pinned optional training dependency, rebuilds and independently
revalidates the corpus, tokenises every example without truncation, and prints the training plan.
Without `--execute`, it creates no Tinker client and makes zero provider calls.

```sh
./scripts/run_training.sh
```

The verified default plan is:

| Item | Value |
| --- | ---: |
| Train examples | 59 |
| Validation examples | 15 |
| Explicit size exclusions | 6 |
| Train tokens | 468,690 |
| Train assistant tokens | 131,352 |
| Maximum train sequence | 33,521 tokens |
| Default optimiser steps | 15 |
| LoRA rank | 32 |
| Learning rate | 0.0001 |
| Batch size / epochs | 4 / 1 |
| Sequence safety limit | 65,536 tokens |

The base model's published context limit is 64K. The safety limit is expressed as 65,536 tokens,
and the longest current example is well below it.

## Explicit paid run

First run the credential-free preflight above. Then place the key in the environment from a private
terminal prompt; do not paste it into chat, a command-line argument or a repository file.

```sh
export TINKER_API_KEY="$(python3 -c 'import getpass; print(getpass.getpass("Tinker API key: "))')"
./scripts/run_training.sh \
  --execute \
  --project-id '<writable-project-id>' \
  --run-dir training/runs/qwen38-sft-v1
unset TINKER_API_KEY
```

Paid execution is refused unless all of `--execute`, `--project-id`, a new `--run-dir` path and a
non-empty `TINKER_API_KEY` are present. A run directory can never be reused, preventing an accidental
second paid run from overwriting earlier local evidence. The default run performs forward-only
validation before and after training, sends only the 59 training examples to `forward_backward` and
the optimiser, pipelines one training request ahead, then saves a resumable state checkpoint and an
indefinite sampler checkpoint. Local metadata includes per-step training NLL, safe numeric provider
metrics, validation NLL before and after, and the validation delta. It never stores the API key.

Lower validation NLL means the checkpoint predicts the development labels better. It is not proof
of spreadsheet correctness. Correctness must be measured by generating workbooks and running the
organiser's evaluator.

## Reading the Tinker session metrics

The retained 400-task benchmark used `SamplingClient`, so it was an inference run rather than a
training run. Its **Average training throughput** and **Training utilization** panels are therefore
expected to be zero; generation and prefill are the relevant rates for that session. The safe
preflight above also leaves training metrics at zero because it deliberately creates no Tinker
client.

The explicit LoRA run submits `forward_backward` and optimiser work, so its training panels should
be populated. Tinker defines training utilisation as the share of each interval spent actively
running forward/backward work; any value above zero shows that training happened, and it need not be
100%. Training tokens per second is an estimated rate over each request's active time. The runner's
one-step lookahead reduces avoidable client-side gaps, while local metadata records token counts,
per-step NLL and numeric provider metrics. Throughput and utilisation describe system activity, not
formula correctness. See Tinker's [session-metrics
definition](https://tinker-docs.thinkingmachines.ai/tinker/session-metrics/).

## Checkpoint evaluation

The paid command prints `SAMPLER_CHECKPOINT=tinker://...`. Use that exact path with the controlled
comparison script:

```sh
export TINKER_API_KEY="$(python3 -c 'import getpass; print(getpass.getpass("Tinker API key: "))')"
export TINKER_PROJECT_ID='<writable-project-id>'
./scripts/evaluate_checkpoint.sh \
  --execute \
  'tinker://<sampler-checkpoint>' \
  training/runs/checkpoint-evaluation-v1
unset TINKER_PROJECT_ID TINKER_API_KEY
```

This is an explicitly paid sampling operation. It independently rebuilds the corpus, derives the
15 validation IDs, and runs those exact tasks twice through FormulaBench's production prompt,
Qwen tool-call parser, workbook writer and safety contracts: once with the base model and once with
the LoRA sampler checkpoint. It then runs the organiser evaluator on both arms and writes a
task-aligned, input-hash-bound `comparison.json`. The output root must be new, and checkpoint mode
cannot resume or mix with a base-model run.

Do not use the values-only `baseline/tinker_predict.py` route to evaluate this checkpoint. That
runner deliberately uses a different prompt and response contract, whereas this SFT experiment is
trained on FormulaBench's production tool-call wire format.

Use the 15 development-validation IDs for checkpoint selection. After the configuration and
checkpoint choice are frozen, compare on the historically named 80-task `held_out` reporting
bucket, then use the 240-task `final_only` reporting bucket for the final reproducibility run.
Neither bucket is statistically untouched: all 400 public tasks already contributed to the
aggregate golden audit and frozen base-model score. Their role here is to keep task-specific labels
out of training and checkpoint selection. Report base-versus-checkpoint pass rate and cell accuracy
together, including every failure and exclusion; do not merge those results into the existing
400-task score.

## Direct commands

The convenience script is equivalent to:

```sh
uv sync --locked --extra native-tinker
uv run python -m training.corpus build \
  --dataset-dir data/spreadsheetbench_verified_400 \
  --split-manifest experiments/public_split.json \
  --out-dir training/generated
uv run python -m training.corpus validate \
  --dataset-dir data/spreadsheetbench_verified_400 \
  --split-manifest experiments/public_split.json \
  --corpus training/generated/corpus.jsonl \
  --manifest training/generated/corpus_manifest.json
uv run --extra native-tinker python -m training.train_lora
```

Inspect all available bounded training arguments with:

```sh
uv run --extra native-tinker python -m training.train_lora --help
```
