# Submission: FormulaBench

## Team

- Team name: FormulaBench
- Member: [`MasteraSnackin`](https://github.com/MasteraSnackin)
- Repository: https://github.com/MasteraSnackin/FormulaBench
- Declared track: Research Track: Excel Formula Generation (SpreadsheetBench)

## Version and evidence status

The default container now runs `formulabench.v2`, an **unscored parity candidate** backed by the
vendored ExactSource package. The vendored source is current at ExactSource commit
`99fe8084bf35a5fca6a2c2e1c9beae802766a618`, and its inference core is the exact core scored in
ExactSource commit `8b84dba1d9263e2123b8f15267239b70ff817907`. This establishes source parity,
not a FormulaBench v2 benchmark result. No same-or-better claim will be made until v2 completes a
fresh 400-task run with the organiser evaluator.

For context, ExactSource's separately published run at that scored commit passed 302/400 tasks
(75.50%) with 80.06% cell accuracy. It used 498 model calls, reported 5,592,930 output tokens and
ran for 6h 39m 36s at the coordinator. Those are upstream source-run figures, not FormulaBench v2
results. The immutable, hash-bound record is
[`full_400_8b84dba.json`](https://github.com/MasteraSnackin/ExactSource/blob/99fe8084bf35a5fca6a2c2e1c9beae802766a618/experiments/full_400_8b84dba.json).

The 133/400 and 33.25% figures below are historical FormulaBench v1 evidence. Its committed
workbooks, traces, prediction manifest, evaluator result and implementation remain intact. The v1
runtime is still available explicitly with `--legacy-engine`.

The canonical v2 image has also passed a credential-free migration audit against all 400 retained
ExactSource outcomes: 369 accepted plans replayed, 31 fallback workbooks matched their inputs byte
for byte, and zero workbook-content mismatches remained. This audit did not call a model or read
golden workbooks. Fresh paid canaries exercised both execution routes: cell task `54513` passed 1/1
target cell through typed operations, and sheet task `23-24` passed 5,510/5,510 target cells through
the restricted Python route. Both used the unchanged organiser evaluator and LibreOffice 26.8.0.3.
The cases were selected as route canaries, so their 2/2 pass count is not a benchmark estimate and
does not alter the unscored status above. The sanitised evidence is in
`experiments/v2_migration_validation.json`.

## What we built and why: historical v1 submission

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

## Current v2 model path

v2 keeps `Qwen/Qwen3.8-27B` fixed and uses the ExactSource reasoning path. It builds up to 48,000
characters of formula-aware workbook context and requests a typed edit plan. Cell-level tasks can
use typed spreadsheet operations only. Sheet-level tasks can use the same operations or a
screened, restricted Python transform. Reasoning is requested for initial calls and ordinary
semantic repairs. Each task has one bounded second-call allowance: either an ordinary semantic
repair or an initial-cell truncation recovery, never both and never a third call. These are runtime
properties, not evidence of spreadsheet correctness.

## Historical v1 model configuration

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

## Historical v1 verified evaluation

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

## Historical v1 run on the 400

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

The current default pipeline runs through:

```sh
python -m formulabench.v2 --dataset-dir=/data --out-dir=/out
```

The Docker runtime reads benchmark inputs from the read-only `/data` mount, writes published
results to `/out`, and confines temporary transform files to its container-local temporary area.
It runs as a non-root user on the pinned Python 3.12.11 slim-bookworm base used by the scored
ExactSource core. Its installed environment, tokenizer cache, application packages and entry point
are root-owned and non-writable by the runtime identity. v2 supports fresh runs and zero-call
preflight. The historical v1 engine is selected explicitly with `--legacy-engine`.

For v2, the wrapper fixes the two mount arguments, mounts the dataset read-only and withholds Tinker
environment variables from preflight. The adapter holds a POSIX advisory lock on the output
directory for the complete run, applies `0700` directory and `0600` file modes, rejects symlinks and
unexpected entry types, and fails before inference when optional serializer dependencies differ
from the canonical image. The screened Python sheet route is defence in depth, not an OS-level
untrusted-code sandbox: its child shares the coordinator's unprivileged container UID. This is
acceptable for the public benchmark but requires stronger isolation and data governance before use
with confidential workbooks.

The remaining runtime details in this section describe v1. Its required credential is
`TINKER_API_KEY`, supplied through the environment. When the account's
Default project is read-only, `TINKER_PROJECT_ID` must identify another writable project. The model,
provider, tokenizer revision, disabled-thinking template, temperature, sample count, configured
retry policy, telemetry policy and response contract cannot be overridden by production arguments. Traces distinguish
the configured model from provider-returned evidence and record the native stop reason, terminal
parse state, accepted response format or sanitised rejection, exact known token counts and the
prompt-plus-parser tool-enforcement method.

An interrupted v1 run is continued with `--legacy-engine --resume`; validated tasks are
skipped and prior `run.log` bytes are preserved. Add `--retry-failures` only when failed
checkpoints should deliberately make another paid attempt. A contract-only write repair can be
applied without credentials or provider calls using:

```sh
./scripts/run_docker.sh \
  data/spreadsheetbench_verified_400 \
  submissions/formulabench \
  --legacy-engine \
  --resume \
  --replay-write-failures
```

## Things to look at

- `formulabench/v2.py` — the default adapter, canonical-runtime guard, fixed CLI contract, output
  lock and private-permission enforcement.
- `exactsource/runner.py` and `exactsource/model.py` — the vendored task coordinator, bounded repair
  policy and Tinker HTTP transport.
- `exactsource/sandbox.py` — screened sheet-transformation worker and its stated isolation limits.
- `tools/verify_v2_parity.py` — credential-free initial-prompt reconstruction and accepted-plan
  replay verifier.
- `experiments/v2_migration_validation.json` — sanitised source, parity and route-canary evidence.
- `formulabench/contract.py`, `formulabench/context.py` and `formulabench/replay.py` — the retained
  historical v1 contract, evidence builder and replay path.
- `formulabench/validate_out.py` — read-only validation of v1 and v2 output artefacts.
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
