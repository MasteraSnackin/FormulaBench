# Submission: FormulaBench

## Team

- Team name: FormulaBench
- Member: [`MasteraSnackin`](https://github.com/MasteraSnackin)
- Repository: https://github.com/MasteraSnackin/FormulaBench
- Declared track: Research Track: Excel Formula Generation (SpreadsheetBench)

## What we built and why

Spreadsheet formula generation often fails while looking plausible: a model can miss the relevant
range, confuse worksheets, omit template cells or return text that never becomes a valid Excel
value. FormulaBench is a deterministic inspect–generate–validate pipeline for the fixed
`Qwen/Qwen3.8-27B` model. The frozen 400-task result uses no fine-tuning or training data. Each task
makes one temperature-zero native Tinker sample with thinking disabled, a pinned Qwen3.8 chat
template and one schema-bound answer tool; the model receives no shell, Python runtime or
unrestricted file tool. FormulaBench reads only the initial workbook in formula and cached-value
modes, builds a sheet-aware evidence map under a 20,000-character prompt budget, and packs answer
ranges before source data. Every output cell carries an exact worksheet and A1 address. Finite
targets must be covered completely, including unchanged cells and blanks, through explicit cells,
compact fills or finite preservation ranges. Typed date objects avoid locale-dependent strings.
Before writing, the contract rejects missing, duplicate and out-of-range cells, explicit external
references, DDE and high-risk executable or external-data functions. Accepted answers are applied
to a fresh workbook copy, reopened for structural and assigned-value round-trip verification, then
published atomically; failures produce a pristine input fallback. Atomic traces and checkpoints make
interruptions and explicit retries auditable without silently repeating paid calls. A separate
replay path can revalidate retained write-failure responses after a contract repair without
credentials or another model call. On the complete public benchmark, the organiser's evaluator
graded all 400 tasks with no missing items or errors. FormulaBench passed 133 tasks: 33.25% pass
rate and 40.56% cell accuracy. That is below the reported approximately 59% organiser baseline.
The configurations differ materially—recommended xhigh reasoning implied by the published runner
and final values there, disabled thinking and live-formula-first strict coverage here—so we report
the gap directly rather than claim an
improvement or a controlled causal result.

## Models

- Inference provider: Tinker
- Base model: `Qwen/Qwen3.8-27B`
- Transport: native Tinker `SamplingClient`, non-streaming token sampling
- Renderer: pinned Qwen3.8 Hugging Face chat template, verified equivalent to
  `qwen3_8_disable_thinking` for the production request
- Thinking: disabled with `enable_thinking=False`
- Reasoning effort: omitted; no effort instruction is injected
- Temperature: `0`
- Samples per task attempt: `1`
- Configured retries: HTTP client `0`; sampling retry wrapper disabled
- Tinker SDK telemetry: forced off before client construction and in the runtime environment
- Fine-tuned checkpoint used for the frozen 400-task result: none
- Fine-tuning data used for the frozen 400-task result: none
- Development-only research checkpoint: Tinker LoRA rank 32 on `Qwen/Qwen3.8-27B`
- Sampler checkpoint:
  `tinker://b2d89255-3d84-5796-84a8-e35f658e3470:train:0/sampler_weights/formulabench-sft-sampler`

### Reproducible training extension

The repository includes an optional post-score research experiment under `training/`. It verifies
the frozen split before opening any golden workbook, permits task-specific goldens only for the 80
development IDs, explicitly excludes six development tasks over 500 target cells, and
deterministically partitions the remaining 74 tasks into 59 training and 15 validation examples.
The generated corpus is bound to the dataset, split, initial workbooks, development goldens,
prompts, answers and rows by SHA-256. The trainer uses the same disabled-thinking Qwen3.8 tool-call
wire format as production, refuses truncation, defaults to zero-call preflight, and requires an
explicit paid-run flag, writable project ID, fresh run directory and privately supplied key.
Its checkpoint comparison uses the same production prompt, tool-call parser, workbook writer and
organiser evaluator for both the base and LoRA arms, then verifies prediction and trace provenance,
rejects all-fallback arms, and records hash-bound task-aligned deltas. A cell-accuracy delta is
withheld whenever evaluator errors would make the underlying denominators non-comparable.

The guarded paid run completed 15 optimiser steps on 59 development examples. Validation NLL fell
from `0.060939` to `0.015557`, and Tinker saved resumable state plus an indefinite sampler
checkpoint. A like-for-like organiser evaluation on the 15 development-validation tasks produced:

| Controlled development-validation comparison | Base | LoRA | Change |
| --- | ---: | ---: | ---: |
| Tasks passed | 7/15 | 9/15 | +2 |
| Pass rate | 46.67% | 60.00% | +13.33 pp |
| Cell accuracy | 74.23% | 77.15% | +2.92 pp |
| Accepted predictions | 14/15 | 15/15 | +1 |
| Evaluator errors (not inference/write failures) | 0 | 0 | 0 |

Three tasks improved from fail to pass while one regressed from pass to fail. This is a real
checkpoint-derived result, but it is deliberately not presented as a replacement 400-task score or
as proof of generalisation. The 15 examples come from the development bucket and were used for
checkpoint selection. The base arm also recorded one fail-closed `workbook_write_failed`
prediction; its untouched fallback workbook remained gradeable, while the checkpoint arm accepted
all 15 predictions. Six development examples with more than 500 target cells were excluded before
the 59/15 split, and each comparison arm used one sample per task. The checkpoint and configuration
must now remain frozen before comparisons
on the historically named 80-task `held_out` and 240-task `final_only` reporting buckets. Those
names define training and selection boundaries, not statistically untouched data: the repository
already reports an aggregate golden audit and a complete 400-task base-model evaluation. The
curated, hash-bound record is
[`experiments/tinker_lora_validation.json`](experiments/tinker_lora_validation.json); see
[`training/README.md`](training/README.md) for the exact protocol. Raw run directories and
recalculated evaluator copies remain local, so the committed record is a transparent hash summary,
not a self-contained or tamper-proof evidence pack.

## Verified evaluation

The first native canary passed development task `54513`: 1/1 task, 1/1 target cell, 100% pass rate
and 100% cell accuracy under the supplied evaluator with LibreOffice recalculation. It is a
development canary, not the 400-task submission score. The model wrote `=C8*(1-E8)`, whose exact
value is 15.7455 and whose existing number format displays 15.75; the golden formula rounds to
15.75 internally, and the supplied evaluator compares both at two-decimal precision. The canary
also exposed a non-fatal Tinker background-poller shutdown warning after its artefacts were
committed. The harder `13-1` run subsequently verified the lifecycle repair with no pending-poller
warning. That initial `13-1` scaffold returned 72 of 120 required cells, so exact validation
rejected every proposed edit for `missing_cells` and published the pristine input. The fallback
scored 82/120 cells, `pass_rate=0.0` and `cell_accuracy=0.6833`; those 82 cells are unchanged
template content, not partially applied model work.

The separately retained revised `13-1` canary included all 120 target cells and all 280 declared
source cells with zero context omissions. It returned all 120 answer cells and passed the strict
contract. With LibreOffice recalculation, the supplied evaluator scored 116/120 cells and
`cell_accuracy=0.9667`. The four wrong grouped amounts mean `pass_rate=0.0`: this is improved cell
accuracy, not a task pass. The run used 6,982 input tokens, 3,113 output tokens and 70,499 ms; its
trace records telemetry disabled and no provider, parser, contract or publication error.

The complete run was then scored with the organiser's evaluator:

```sh
uv run python evaluate.py \
  --predictions=submissions/formulabench/predictions.jsonl \
  --all \
  --out=submissions/formulabench/results.json
```

No partial-run or alternative-evaluator result will be presented as the submission score.

## Our run on the 400

The completed run is retained at `submissions/formulabench/`:

- `submissions/formulabench/predictions.jsonl`
- `submissions/formulabench/outputs/`
- `submissions/formulabench/traces/`
- `submissions/formulabench/run.log`
- `submissions/formulabench/results.json`

The first pass produced 350 accepted responses and 50 fail-closed fallbacks. A deterministic,
credential-free replay then revalidated seven retained `workbook_write_failed` responses under a
corrected exact round-trip contract. Four were published successfully, three remained failures,
and the replay made zero additional model calls. The final set contains 354 accepted workbooks and
46 pristine input fallbacks.

LibreOffice `7.4.7.2` recalculation and the organiser's evaluator graded all 400 tasks with zero
missing items and zero errors. This public self-evaluation result is 133/400 task passes,
`pass_rate=0.3325` and
`cell_accuracy=0.4056`; cell-level task pass rate is `0.3273` and sheet-level task pass rate is
`0.3440`. The complete result is `submissions/formulabench/results.json`, SHA-256
`c1e6fb6b540bb7272f4c537773e2a4826dd649d882c9b7fb549df9bccc66be0d`.

This is below the organiser's reported approximately 59% reference. The published organiser
runner implies the cookbook-recommended `qwen3_8_xhigh_reasoning` renderer and requests final
values, while FormulaBench fixes thinking off and prioritises formulas under a stricter target and
workbook contract. The organisers did not publish the reference run's traces, so the renderer is
inferred from their code and locked cookbook behaviour, not independently verified run metadata.
The score gap therefore cannot be attributed to one component.

A pre-recalculation static comparison found 31,240 authored or changed formulas and no newly
introduced formula-gate, defined-name, macro, connection, sensitive-part or external-relationship
problem. The inputs already contained 92 sensitive OOXML parts and 153 external relationships, so
evaluation ran in a networkless, read-only-root container with all Linux capabilities dropped and
no host credentials.

The immutable `canary-54513-native-2`, `canary-13-1-native` and `canary-13-1-native-v2`
directories remain separate local development evidence. Their sanitised facts are recorded in
`experiments/live_preflight.json`; their gitignored raw traces are not represented as the final
submission run.

## Code

The submitted pipeline runs through:

```sh
python -m formulabench.capture --dataset-dir=/data --out-dir=/out
```

The Docker runtime reads only from `/data` and writes to `/out`. It runs as a non-root user.
The required credential is `TINKER_API_KEY`, supplied through the environment. When the account's
Default project is read-only, `TINKER_PROJECT_ID` must identify another writable project. The model,
provider, tokenizer revision, disabled-thinking template, temperature, sample count, configured
retry policy, telemetry policy and response contract cannot be overridden by production arguments. Traces distinguish
the configured model from provider-returned evidence and record the native stop reason, terminal
parse state, accepted response format or sanitised rejection, exact known token counts and the
prompt-plus-parser tool-enforcement method.

An interrupted run is continued with the same command plus `--resume`; validated tasks are
skipped and prior `run.log` bytes are preserved. Add `--retry-failures` only when failed
checkpoints should deliberately make another paid attempt. A contract-only write repair can be
applied without credentials or provider calls using:

```sh
./scripts/run_docker.sh \
  data/spreadsheetbench_verified_400 \
  submissions/formulabench \
  --resume \
  --replay-write-failures
```

## Things to look at

- `formulabench/contract.py` — sheet-qualified targets, strict validation and atomic writes.
- `formulabench/context.py` — formula-aware, bounded and spatially distributed workbook evidence.
- `formulabench/replay.py` — hash-bound, credential-free replay of eligible stored responses.
- `formulabench/validate_out.py` — read-only validation of final submission artefacts.
- `tests/` — synthetic and all-initial-workbook compatibility coverage.
- `experiments/public_input_audit.json` — reproducible target-size, crop and formula-cache
  measurements from public initial workbooks only.
- `experiments/public_split.json` — frozen, deterministic 80/80/240 public evaluation split,
  cryptographically bound to its manifest and initial workbooks and produced without opening
  golden workbooks.
- `experiments/public_benchmark_audit.json` — a later aggregate structural audit that did open all
  public golden workbooks; it publishes no task-specific answer values and makes clear why the
  public buckets are reporting partitions rather than statistically untouched holdouts.
- `experiments/tinker_lora_validation.json` — curated training and like-for-like checkpoint
  evaluation evidence, including configuration, summary metrics and cryptographic bindings.
- `PROVENANCE.md` — exact organiser scaffold revision and dataset licence boundary.

## Public submission links

- Presentation: https://masterasnackin.github.io/FormulaBench/presentation.html
- Demo video: https://masterasnackin.github.io/FormulaBench/video.html
- Live demo: https://masterasnackin.github.io/FormulaBench/

The live demo is a static, evidence-backed results viewer. It makes no paid model calls, requires
no API keys and reports the complete 400-task public-benchmark result alongside the method,
failure analysis and reproducibility evidence.
