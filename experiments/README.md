# Evaluation plan

FormulaBench treats the organiser's evaluator as the only source of benchmark scores. Inference
opens `dataset.json` and each initial workbook, never a golden file. The input-only audit and frozen
split below also did not open golden workbooks. A later, explicitly separate aggregate benchmark
audit did open all 400 public golden workbooks to measure structural properties; it publishes no
task-specific answer values, but its findings informed general contract design. The public split
must therefore be treated as reporting buckets rather than statistically untouched holdouts.

## Reproduce the input audit

```sh
uv run python experiments/analyse_public_inputs.py \
  --dataset-dir data/spreadsheetbench_verified_400 \
  > /tmp/formulabench-public-input-audit.json
diff -u experiments/public_input_audit.json /tmp/formulabench-public-input-audit.json
```

The audit motivated three changes that can be tested independently: retaining formula text and
spatial evidence, qualifying every answer by worksheet, and expanding validated compact fills
for large repeated-formula targets.

## Reproduce the aggregate public benchmark audit

```sh
uv run python experiments/analyse_public_benchmark.py \
  --dataset-dir data/spreadsheetbench_verified_400 \
  > /tmp/formulabench-public-benchmark-audit.json
diff -u experiments/public_benchmark_audit.json \
  /tmp/formulabench-public-benchmark-audit.json
```

This later audit reads both initial and golden public workbooks. It reports only aggregate
structure and an untouched-input baseline; it emits no task IDs, instructions, cell coordinates or
cell values. Its existence does not alter the earlier split membership or hashes, but aggregate
knowledge did influence the general response contract. In particular, 88 of 125 sheet-level tasks
contain required blank cells, 86 contain correct nonblank prefill but still require changes, and 43
require clearing 100,752 existing nonblank target cells. Across both instruction levels, 37 tasks
contain 10,319 date/datetime target cells. These measurements motivated exact complete coverage,
explicit preservation and typed Excel date values.

## Reproduce the frozen public split

The public task split was generated before the first credentialled model run. It uses only the
manifest and validated initial workbooks; it does not locate or open golden workbooks.

```sh
uv run python experiments/make_public_split.py \
  --dataset-dir data/spreadsheetbench_verified_400 \
  > /tmp/formulabench-public-split.json
diff -u experiments/public_split.json /tmp/formulabench-public-split.json
```

`public_split.json` contains 80 development IDs, 80 historically named `held_out` IDs and 240
historically named `final_only` IDs. The allocation is deterministic and stratified by the
manifest's instruction type and an initial-workbook-resolved target-size proxy. The three lists
are disjoint, exhaustive and kept in manifest order. The artefact records the manifest SHA-256 and
an ordered aggregate SHA-256 over every task ID and exact initial-workbook file, so both declared
split inputs are bound. The later aggregate golden audit changes how the latter two lanes are
interpreted, not how or when the split was constructed.

## Credentialed evaluation sequence

Sanitised request and outcome facts are retained in
[`live_preflight.json`](live_preflight.json). Compatibility-endpoint isolation established that
explicit `xhigh` was rejected and that this model/endpoint combination could not reliably return a
complete structured answer with thinking disabled. The fixed submission therefore uses the native
Tinker client, a pinned Qwen3.8 thinking-disabled chat template, temperature zero, one sample and
no FormulaBench-level automatic retry.

After selecting a writable Tinker project, native development task `54513` completed in 4,730 ms
with 6,808 input and 57 output tokens. The sole terminal tool call wrote `=C8*(1-E8)` to
`Sheet1!F8`; LibreOffice recalculation and the supplied evaluator returned 1/1 correct cells,
`pass_rate=1.0` and `cell_accuracy=1.0`. Its immutable log exposed a non-fatal Tinker futures-poller
warning after publication. The lifecycle repair was then verified by the initial native `13-1`
run, which produced no pending-task warning.

That harder `13-1` response returned 72 of 120 required cells. Exact validation rejected all edits
for `missing_cells` and published the pristine input; the fallback scored 82/120 cells,
`pass_rate=0.0` and `cell_accuracy=0.6833`. Those 82 cells are the unchanged-input signature, not
partially applied model output. Failure analysis led to answer-first compact target rows, strict
typed Excel dates/datetimes, explicit preservation ranges and stronger aggregation instructions.
The revised prompt contains all 120 target cells and 280 declared source cells with zero
omissions. Its separate live canary returned all 120 cells and passed the output contract. The
supplied evaluator with LibreOffice recalculation scored 116/120 cells,
`cell_accuracy=0.9667`, but `pass_rate=0.0` because four amount cells in the final grouped section
were wrong. It used 6,982 input tokens, 3,113 output tokens and 70,499 ms. This is development
evidence, not the complete benchmark score reported below.

1. Iterate on task-specific behaviour only within the frozen 80-task development bucket, and
   preserve every attempted configuration, rejection and official-evaluator result.
2. Use the remaining public buckets for transparent validation and final reporting. Do not label
   them statistically untouched: aggregate golden-workbook findings informed general design.
3. Freeze model, prompt, contract and runtime before the organiser's private evaluation. The
   private dataset remains the genuine unseen test.
4. Score each complete public run with the organiser's evaluator and retain its full
   `results.json`, workbooks, traces and `run.log`.

## Complete public run

The fixed thinking-disabled FormulaBench configuration ran once across all 400 public tasks. Its
first pass accepted 350 responses and failed closed on 50. A later deterministic replay considered
only seven retained responses with the exact `workbook_write_failed` failure code after the local
round-trip contract was corrected. It published four, left three unchanged and made zero model
calls. The final output therefore contains 354 accepted workbooks and 46 pristine input fallbacks:
26 `model_call_failed`, 12 `missing_cells`, three `duplicate_cells`, three
`workbook_write_failed`, one `invalid_response_schema` and one `extra_unmerge_ranges`.

The official evaluator ran after LibreOffice `7.4.7.2` recalculation and graded all 400 tasks with
zero missing items and zero errors. The retained summary is:

| Metric | Result |
| --- | ---: |
| Tasks passed | 133/400 |
| Pass rate | 0.3325 |
| Cell accuracy | 0.4056 |
| Cell-level task pass rate | 0.3273 |
| Sheet-level task pass rate | 0.3440 |

The result is [`../submissions/formulabench/results.json`](../submissions/formulabench/results.json),
SHA-256 `c1e6fb6b540bb7272f4c537773e2a4826dd649d882c9b7fb549df9bccc66be0d`.
Recalculation was isolated in a container with networking disabled, a read-only root filesystem,
all Linux capabilities dropped and no host credentials. A uniform pre-recalculation comparison
found 31,240 authored or changed formulas and no newly introduced formula-gate, defined-name,
macro, connection, sensitive-part or external-relationship problem. It separately counted 92
sensitive OOXML parts and 153 external relationships inherited from the supplied inputs.

The 33.25% pass rate is below the organiser's reported approximately 59% Qwen3.8-27B reference,
but the two configurations are materially different. The published organiser runner implies the
cookbook-recommended `qwen3_8_xhigh_reasoning` renderer and requests final scalar values.
FormulaBench disables thinking, prioritises executable formulas, uses a bounded 20,000-character
context and rejects incomplete target partitions. The organisers did not publish the reference
run's traces, so its renderer is inferred from the published runner and locked cookbook behaviour,
not independently verified run metadata. These results support comparison of complete systems,
not attribution of the gap to one component.

Reproduce the zero-call replay with:

```sh
./scripts/run_docker.sh \
  data/spreadsheetbench_verified_400 \
  submissions/formulabench \
  --resume \
  --replay-write-failures
```

The replay path revalidates input, trace and stored-response hashes, is mutually exclusive with
paid retry, and annotates successful traces with `additional_model_calls=0`.

## Failure analysis

The retained results, predictions and traces show that strict rejection is not the main source of
the score gap. The 46 fallback workbooks account for 9,143 of 177,058 wrong cells, or 5.16%.
The 221 contract-valid but semantically incorrect workbooks account for the remaining 167,915
wrong cells, or 94.84%. Even if all 46 rejected tasks had become passes, the result would have
been at most 179/400, or 44.75%, before considering whether their answers were correct.

| Target cells | Tasks | Passes | Pass rate | Wrong cells |
| --- | ---: | ---: | ---: | ---: |
| 1 | 40 | 24 | 60.00% | 16 |
| 2–10 | 123 | 37 | 30.08% | 291 |
| 11–100 | 175 | 59 | 33.71% | 2,213 |
| 101–500 | 35 | 11 | 31.43% | 2,406 |
| 501–1,000 | 10 | 0 | 0.00% | 3,703 |
| 1,001–10,000 | 13 | 2 | 15.38% | 18,353 |
| Over 10,000 | 4 | 0 | 0.00% | 150,076 |

The four largest transformations therefore contain 84.76% of every wrong cell. Asking the model
to enumerate these outputs is the clearest cell-accuracy bottleneck. The 400 traces also show a
strong association with context truncation: 47/195 truncated tasks passed (24.10%), compared with
86/205 non-truncated tasks (41.95%). This is not proof of causation because complex workbooks are
more likely to be truncated, but it justifies task-aware retrieval rather than uniform sampling.

The 26 provider-response failures comprise 14 responses that exhausted the 8,192-token limit,
nine invalid tool-call envelopes and three invalid JSON envelopes. The remaining 20 failures were
local contract or workbook rejections. Separately, the evaluator's limited first-five mismatch
samples expose Excel runtime errors in 27 accepted workbooks; this is a lower bound rather than a
complete error-cell count.

The evidence supports two routes for a next version: formula synthesis plus local recalculation
checks for cell-level formula tasks, and a small deterministic operation executor for large
filter, append, sort, group, deduplication, transpose, clear and string-transformation tasks. That
would let the model select an operation and its parameters without calculating or serialising
tens of thousands of cells. No such post-score redesign is represented as part of the frozen run.

## Optional native Tinker and fine-tuning experiments

`baseline/tinker_predict.py` uses [Tinker Cookbook's Qwen3.8
renderer](https://github.com/thinking-machines-lab/tinker-cookbook/blob/main/tinker_cookbook/renderers/qwen3_8.py)
and native sampling client. It defaults to the renderer recommended for the base model and accepts
an explicit `--renderer` for xhigh, medium, low or thinking-disabled ablations. It provides a
controlled comparison if the Anthropic-compatible transport remains unreliable, and it can sample
a `tinker://...` LoRA checkpoint. Keep this dependency in the optional `native-tinker` environment
rather than the production image. This remains a values-only baseline; it is not the correct
evaluation wire format for a checkpoint trained on FormulaBench's formula-and-tool-call responses.

The implemented SFT path is documented in [`training/README.md`](../training/README.md). Its
reproducible corpus contains 59 training examples and 15 development-validation
examples, with six explicit over-500-cell exclusions, 468,690 total train tokens and 15 planned
optimiser steps. The preflight verifies that Cookbook and production prompt tokens are identical
and makes zero provider calls. The paid rank-32 LoRA run completed all 15 steps and reduced
development-validation NLL from `0.060939` to `0.015557`. The like-for-like organiser evaluation
then improved pass rate from 46.67% (7/15) to 60.00% (9/15) and cell accuracy from 74.23% to
77.15%, with zero evaluator errors in either arm. Three tasks changed from fail to pass and one
regressed from pass to fail. The base arm accepted 14/15 predictions and recorded one fail-closed
`workbook_write_failed`; its unchanged fallback was still graded. The checkpoint accepted all 15.
The curated evidence is
[`tinker_lora_validation.json`](tinker_lora_validation.json), including its source hashes and
checkpoint provenance. This is checkpoint-selection evidence on a small development split, not a
replacement for the frozen 400-task result or proof of broader generalisation. The raw run and
recalculated evaluator artefacts are not committed, so the public record does not independently
reproduce every recorded hash.

Only the frozen development bucket may supply task-specific prompt examples, supervised examples,
reward examples or checkpoint-selection cases. A spreadsheet-specific training example should
pair the same bounded workbook prompt used at inference with a validated sheet-qualified JSON
answer. The organiser's workbook evaluator, not the cookbook's general-language benchmarks,
remains the correctness oracle. Aggregate public benchmark characteristics may guide general
contract design, but no task-specific golden value outside development may be copied into prompts,
runtime, training data or hand-tuned task logic.

Compare `pass_rate` first and `cell_accuracy` second. Also report cell-level versus sheet-level
pass rate, structural failure codes, target-size buckets (`1`, `2–10`, `11–100`, `101–500`,
`501–1000`, `1001+`) and whether the context was truncated. Do not relabel a partial run, oracle
check or structurally valid workbook as a correctness result.

The private dataset remains genuinely held out: freeze the inference code and configuration
before submission, and report any mismatch between public and private performance as a limit
rather than retuning to private examples.
