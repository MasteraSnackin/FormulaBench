# FormulaBench

A reproducible, fail-closed pipeline for generating and validating Excel formulas.

[![Checks](https://github.com/MasteraSnackin/FormulaBench/actions/workflows/ci.yml/badge.svg)](https://github.com/MasteraSnackin/FormulaBench/actions/workflows/ci.yml)
![Submitted Python runtime](https://img.shields.io/badge/Python-3.11%20runtime-3776AB?logo=python&logoColor=white)
[![Research Track: Excel Formula Generation](https://img.shields.io/badge/Research%20Track-Excel%20Formula%20Generation-1F6FEB)](SUBMISSION.md)

## Description

FormulaBench takes a plain-English spreadsheet instruction and an initial Excel workbook, then
attempts to produce the requested `.xlsx` workbook using `Qwen/Qwen3.8-27B` through Tinker. It
serves formula-generation researchers and spreadsheet analysts who need generated changes to
remain reviewable. It addresses model outputs that look plausible while missing required cells,
writing to the wrong sheet or introducing unsafe formulas.

The pipeline builds sheet-aware evidence, requests one structured model response, validates exact
target coverage and formula safety, then publishes either a validated edit or a pristine input
fallback. It retains one workbook and trace file per task, plus a prediction record in the shared
checkpoint manifest. A model attempt adds the prompt, any available parsed response, token counts,
latency and failure state to that trace.

This repository is the **Research Track: Excel Formula Generation (SpreadsheetBench)** entry for
the Encode x Ylookup Rebuild Private Markets Hackathon. The frozen 400-task result uses prompt
engineering and deterministic validation rather than fine-tuning. The repository now also contains
a separate, development-only Tinker LoRA experiment with a controlled 15-task validation result.
That small comparison is reported separately and does not replace or alter the frozen score below.

### Public self-evaluation result

| SpreadsheetBench Verified self-evaluation | Result |
| --- | ---: |
| Tasks graded | 400/400 |
| Tasks passed | 133/400 |
| Pass rate | 33.25% |
| Cell accuracy | 40.56% |
| Cell-level task pass rate | 32.73% |
| Sheet-level task pass rate | 34.40% |
| Missing tasks / evaluator errors | 0 / 0 |

The complete self-evaluation result is in
[`submissions/formulabench/results.json`](submissions/formulabench/results.json), SHA-256
`c1e6fb6b540bb7272f4c537773e2a4826dd649d882c9b7fb549df9bccc66be0d`. A local run with the
organiser-supplied evaluator reported 133 passes from 400 tasks. The
[organiser's Research Track README](https://github.com/ylookup/encode-hackathon/blob/37d9016264762a25cae49e077cd0893055bd9093/research/README.md#L88)
lists a 59.0% one-shot, values-only Qwen3.8-27B reference. FormulaBench disables thinking,
prioritises formulas and applies a stricter fail-closed contract, so the results are not directly
comparable.

### Controlled Tinker LoRA validation result

FormulaBench also trained a rank-32 LoRA for `Qwen/Qwen3.8-27B` through Tinker on 59 verified
development examples. The run completed all 15 optimiser steps, saved an indefinite sampler
checkpoint, and reduced development-validation NLL from `0.060939` to `0.015557`. The same 15
development-validation tasks were then run through the base model and checkpoint using identical
production prompts, sampling settings, parser, workbook writer and organiser evaluator.

| 15-task development-validation comparison | Base model | LoRA checkpoint | Change |
| --- | ---: | ---: | ---: |
| Tasks passed | 7/15 | 9/15 | +2 tasks |
| Pass rate | 46.67% | 60.00% | +13.33 pp |
| Cell accuracy | 74.23% | 77.15% | +2.92 pp |
| Cell-level task pass rate | 45.45% | 63.64% | +18.19 pp |
| Sheet-level task pass rate | 50.00% | 50.00% | 0.00 pp |
| Accepted predictions | 14/15 | 15/15 | +1 |
| Evaluator errors (not inference/write failures) | 0 | 0 | 0 |

Three tasks changed from fail to pass, one changed from pass to fail, and the remaining eleven kept
their whole-task outcome. The base arm had one fail-closed `workbook_write_failed` prediction; its
unchanged fallback was still graded, while the checkpoint arm accepted all 15 predictions. The
result was locally validated against prediction, trace and workbook hashes, with its curated record
in
[`experiments/tinker_lora_validation.json`](experiments/tinker_lora_validation.json). It is useful
checkpoint-selection evidence, not a new 400-task benchmark score: the split is small, drawn from
the development bucket and used to select this checkpoint, uses one sample per task and excludes
six development examples whose targets exceed 500 cells. A broader post-freeze comparison is still
required before claiming generalisation. Raw run directories and recalculated evaluator workbooks
are not committed, so the public JSON is a hash-recorded summary rather than a self-contained proof.

### A pass and a near miss

| Evaluator outcome | Task | What FormulaBench had to do | What happened |
| --- | --- | --- | --- |
| Fully passed | [`54513`](submissions/formulabench/traces/54513.jsonl) | Calculate the price of a $34.99 item after a 55% discount. | FormulaBench wrote `=C8*(1-E8)` to `Sheet1!F8`. It calculated `15.7455`, displayed as `$15.75`, and the evaluator marked the target cell correct. |
| Failed, but nearly complete | [`13-1`](submissions/formulabench/traces/13-1.jsonl) | Combine data from one sheet into another, merge duplicate rows by date and reference, sort the results and calculate totals. | FormulaBench completed 116 of 120 target cells correctly. Four amount cells were wrong, so the strict whole-task result was a failure. |

The four mismatches in task `13-1` were:

| Cell | Expected | FormulaBench produced |
| --- | ---: | ---: |
| `LISTS!D25` | 1,990 | 2,020 |
| `LISTS!D30` | 534 | 394 |
| `LISTS!D31` | 401 | 276 |
| `LISTS!D32` | 6,495 | 6,260 |

That task reached 96.67% cell accuracy, but the final total was among the four incorrect cells. The
whole task therefore failed. This is why FormulaBench evaluates the recalculated workbook rather
than treating a formula that runs as a correct answer.

### Judge resources

- [Submission overview](SUBMISSION.md)
- [System architecture](ARCHITECTURE.md)
- [Live research overview](https://masterasnackin.github.io/FormulaBench/)
- [Demo video](https://masterasnackin.github.io/FormulaBench/video.html)
- [Presentation viewer](https://masterasnackin.github.io/FormulaBench/presentation.html)
- [Self-evaluation results](submissions/formulabench/results.json)
- [Tinker LoRA validation evidence](experiments/tinker_lora_validation.json)
- [Prediction manifest](submissions/formulabench/predictions.jsonl)
- [Generated workbooks](submissions/formulabench/outputs/)
- [Model traces](submissions/formulabench/traces/)
- [Evaluation and failure analysis](experiments/README.md)
- [Dataset and scaffold provenance](PROVENANCE.md)

## Table of Contents

- [Description](#description)
- [Features](#features)
- [Tech Stack](#tech-stack)
- [Architecture Overview](#architecture-overview)
- [Installation](#installation)
- [Usage](#usage)
- [Configuration](#configuration)
- [Screenshots / Demo](#screenshots--demo)
- [API / CLI Reference](#api--cli-reference)
- [Tests](#tests)
- [Roadmap](#roadmap)
- [Contributing](#contributing)
- [License](#license)
- [Contact / Support](#contact--support)

## Features

- Sheet-qualified targets and formula-aware workbook context from formula and cached-value views.
- A deterministic 20,000-character prompt budget with complete compact ranges and spatial sampling.
- Strict typed responses with exact validation of missing, duplicate and out-of-target cells.
- A static formula safety gate followed by atomic write and reopen checks.
- Pristine input fallbacks when context construction, inference, validation or writing fails.
- Validated resume, explicit failure retry and credential-free replay of eligible stored responses.
- Per-task workbook and trace files, with one prediction record per task in an atomically rewritten
  shared checkpoint manifest.
  A failure before a model call can leave the required trace file empty.
- Read-only output validation with optional scanning for an explicitly configured secret value.
- A deterministic development-only SFT corpus, guarded LoRA trainer and completed checkpoint
  comparison with split, workbook, prompt, answer, prediction, trace and output hashes.
- A static GitHub Pages viewer for the retained result, demo video and presentation.

## Tech Stack

| Area | Technology |
| --- | --- |
| Language | Python. The submitted image uses 3.11. `pyproject.toml` declares `>=3.11`. |
| Workbook processing | `openpyxl==3.1.5` |
| Validation | Pydantic 2 |
| Inference | Tinker SDK `0.27.1` with `Qwen/Qwen3.8-27B` |
| Optional training | Tinker Cookbook `0.5.7`, LoRA SFT and held-back development validation |
| Prompt rendering | Transformers chat template and Tokenizers |
| Dependency management | `uv` with a committed lockfile |
| Evaluation | Organiser evaluator and LibreOffice Calc |
| Packaging and runtime | Hatchling and Docker |
| Tests and code quality | pytest and Ruff |
| Continuous integration | GitHub Actions on Ubuntu 24.04 |
| Evidence site | Static HTML, CSS and JavaScript deployed through GitHub Pages |
| Persistence | Local `.xlsx`, JSONL, JSON and log files. No database. |

## Architecture Overview

```mermaid
flowchart TB
    operator["Operator"] --> system["FormulaBench CLI and task runner"]
    inputs[("Manifest and initial workbooks")] --> system
    system -.->|"Sample request"| tinker["Tinker with Qwen3.8-27B"]
    tinker -.->|"Token response"| system
    system --> contract["Response and workbook contract"]
    system -->|"Context or provider failure"| fallback["Pristine workbook fallback"]
    contract -->|"Contract valid"| artefacts[("Workbooks and audit artefacts")]
    contract -->|"Rejected or write failure"| fallback
    fallback --> artefacts
    system -->|"Predictions, traces and log"| artefacts
    artefacts --> validator["Read-only validator"]
    operator -.->|"Separate command"| evaluator["LibreOffice and organiser evaluator"]
    artefacts -->|"Workbooks and predictions"| evaluator
    goldens[("Golden workbooks")] --> evaluator
    evaluator --> results[("Self-evaluation result")]
```

`formulabench.capture` supervises the production CLI, which loads each task and builds formula-aware
workbook context. For tasks that reach the provider, the runner exchanges one logical sample with
Tinker, applies only responses that pass the contract and uses a pristine fallback for failures,
while the capture supervisor owns the process log. Complete runs validate automatically,
golden-aware scoring stays separate, and the read-only GitHub Pages viewer publishes frozen evidence
outside the inference path, as detailed in [ARCHITECTURE.md](ARCHITECTURE.md).

## Installation

### Requirements

- Git.
- Python 3.11 for parity with the submitted image. The project declares `>=3.11`, but CI does not
  test a multi-version matrix.
- [`uv`](https://docs.astral.sh/uv/).
- Docker for the submitted one-command runtime.
- LibreOffice Calc to rerun recalculation-based scoring locally. Exact parity across LibreOffice
  releases is not established.
- A Tinker account and API key for ordinary inference. The direct Python CLI requires the key only
  while work is pending. The Docker wrapper requires it for every mode except preflight and
  stored-response replay, including a fully completed resume. Tests and validation do not need it.

### Set up from scratch

```sh
git clone https://github.com/MasteraSnackin/FormulaBench.git
cd FormulaBench
uv sync --locked
uv run python data/download.py
```

On a fresh download, the script checks the SpreadsheetBench Verified archive against its pinned
SHA-256 before extraction. The dataset directory is intentionally excluded from Git.

Optional comparison environments are kept out of the production image:

```sh
# Organiser-style OpenRouter comparison dependencies
uv sync --locked --extra baseline

# Native Tinker Cookbook and checkpoint experiments
uv sync --locked --extra native-tinker
```

### Reproduce the safe Tinker training preflight

The training script rebuilds the frozen development-only corpus, verifies it independently,
tokenises all 74 eligible examples and prints the bounded plan. Its default mode makes zero Tinker
provider calls:

```sh
./scripts/run_training.sh
```

Paid training is a separate, explicit action requiring `--execute`, a writable Tinker project, a
local run directory and a key supplied privately through the environment. See the complete
[training protocol](training/README.md), including leakage boundaries, measured token counts,
validation NLL and production-parity checkpoint evaluation. The completed experiment improved the
15-task development-validation pass rate from 46.67% to 60.00% and cell accuracy from 74.23% to
77.15%. The separate evaluation script scored both arms with the organiser evaluator and emitted a
provenance-and-hash-bound comparison. It records accepted and failed predictions, rejects an
all-fallback arm, and does not present cell-accuracy deltas as comparable when evaluator errors
changed the denominator. The wrappers expose the API key only to paid trainer or inference
processes. Generated corpus and raw run directories remain ignored by Git; the curated result is in
[`experiments/tinker_lora_validation.json`](experiments/tinker_lora_validation.json) and does not
alter the retained 400-task submission evidence.

## Usage

### Run the submitted Docker workflow

Export `TINKER_API_KEY` in your private shell rather than storing it in the repository. Set
`TINKER_PROJECT_ID` only if the account's default project is read-only. A fresh run requires an
empty output directory and may make Tinker requests that consume credits or incur cost.

```sh
export TINKER_API_KEY="$(python3 -c 'import getpass; print(getpass.getpass("Tinker API key: "))')"
./scripts/run_docker.sh \
  data/spreadsheetbench_verified_400 \
  submissions/my-run
```

The wrapper builds `formulabench:latest`, mounts the dataset read-only at `/data`, mounts the
selected output directory at `/out`, and runs as the current host user.

### Run directly with Python instead

```sh
uv run python -m formulabench.capture \
  --dataset-dir=data/spreadsheetbench_verified_400 \
  --out-dir=submissions/my-python-run
```

Use a known development task for a bounded canary:

```sh
uv run python -m formulabench.capture \
  --dataset-dir=data/spreadsheetbench_verified_400 \
  --out-dir=submissions/canary-54513 \
  --ids=54513
```

### Resume or retry

Ordinary resume reuses only validated checkpoints and does not retry recorded failures:

```sh
./scripts/run_docker.sh \
  data/spreadsheetbench_verified_400 \
  submissions/my-run \
  --resume
```

Reschedule validated failures only when another inference attempt is intentional. A task consumes
credits or incurs cost only if it reaches the provider:

```sh
./scripts/run_docker.sh \
  data/spreadsheetbench_verified_400 \
  submissions/my-run \
  --resume \
  --retry-failures
```

Replay eligible stored `workbook_write_failed` responses after a contract repair, without a
credential or provider call:

```sh
./scripts/run_docker.sh \
  data/spreadsheetbench_verified_400 \
  submissions/my-run \
  --resume \
  --replay-write-failures
```

`--retry-failures` and `--replay-write-failures` are mutually exclusive.

### Validate and score outputs

```sh
uv run python -m formulabench.validate_out \
  --dataset-dir=data/spreadsheetbench_verified_400 \
  --out-dir=submissions/my-run \
  --expected-model=Qwen/Qwen3.8-27B \
  --secret-env=TINKER_API_KEY

uv run python evaluate.py \
  --predictions=submissions/my-run/predictions.jsonl \
  --all \
  --out=submissions/my-run-results.json
```

Keep the score file outside a run that may be resumed. The resume allow-list rejects additional
files at the output root, including `results.json`.

The inference output directory has this stable layout:

```text
submissions/my-run/
├── predictions.jsonl
├── run.log
├── outputs/
│   └── <task-id>.xlsx
└── traces/
    └── <task-id>.jsonl
```

## Configuration

### Production and evaluation environment variables

| Variable | Required | Purpose |
| --- | --- | --- |
| `TINKER_API_KEY` | Docker wrapper except preflight/replay. Direct CLI when inference is pending. | Authenticates the native Tinker sampling client. |
| `TINKER_PROJECT_ID` | Account-dependent | Selects a writable project when the account default is read-only. |
| `FORMULABENCH_IMAGE` | No | Overrides the Docker image name used by `scripts/run_docker.sh`. |
| `SOFFICE` | No | Selects the LibreOffice executable used only by the separate evaluator. |

Do not commit credentials. Tinker telemetry is forced off before client construction and in the
container environment.

### Fixed inference settings

| Setting | Value |
| --- | --- |
| Model | `Qwen/Qwen3.8-27B` |
| Transport | Native Tinker sampling client |
| Renderer | Pinned Qwen3.8 thinking-disabled chat template |
| Tokenizer revision | `1d4bf0f2ff6012fd82039f2fa52739d0dd7c60c0` |
| Thinking | Disabled |
| Temperature | `0` |
| Samples per task attempt | `1` |
| Maximum output tokens | `8,192` |
| Maximum authored prompt | `20,000` characters |
| Provider HTTP retries | `0` |
| Default maximum sampling concurrency | `4` |

Production model settings are centralised in [`formulabench/constants.py`](formulabench/constants.py)
and are not exposed as CLI overrides. The application and runtime dependency resolution is locked
in [`uv.lock`](uv.lock). [`Dockerfile`](Dockerfile) defines the container build and runtime.

Each attempt that reaches the provider requests one logical sample and disables FormulaBench's own
HTTP and sampling retries. The Tinker SDK may still retry idempotent submission or retrieval
operations for that same sample.

## Screenshots / Demo

The FormulaBench inference engine is CLI-based. Its hosted demo is a read-only evidence viewer, not
an inference interface. These panels are built from the retained workbook, process log and result
from the organiser-supplied evaluator, without making a fresh model call.

### Generated workbook

![FormulaBench workbook evidence for passing public task 54513](docs/images/readme-workbook-task-54513.png)

Task `54513` produced `=C8*(1-E8)` in `Sheet1!F8`. The retained workbook displays `15.75`, and the
organiser-supplied evaluator marked its single target cell correct.

### Complete public run

![FormulaBench retained run and evaluation evidence](docs/images/readme-public-run-evidence.png)

The panel combines exact excerpts from `submissions/formulabench/run.log` with figures derived from
`submissions/formulabench/results.json`. It records the 400-task run, credential-free replay and
final 33.25% task pass rate. The replay made no additional provider calls.

### Demo and presentation

- [Open the live research overview](https://masterasnackin.github.io/FormulaBench/).
- [Watch the 67.4-second demo video](https://masterasnackin.github.io/FormulaBench/video.html).
- [View the nine-slide presentation](https://masterasnackin.github.io/FormulaBench/presentation.html),
  [open the PDF](presentation/FormulaBench-Hackathon-Deck.pdf), or
  [download the PowerPoint source](presentation/FormulaBench-Hackathon-Deck.pptx).

GitHub Pages serves this static evidence surface. It cannot upload a workbook or invoke Tinker.

<details>
<summary>Accessible run transcript</summary>

```text
# Initial run
FormulaBench  model=Qwen/Qwen3.8-27B  tasks=400  remaining=400  concurrency=4
completed  ok=350  failed=50
output validation  ok=True  issues=0  workbooks=400/400

# Stored-response replay and final validation
stored-response replay  eligible=7  replayed=4  unchanged=3  additional_model_calls=0
completed  ok=354  failed=46
output validation  ok=True  issues=0  workbooks=400/400
```

</details>

## API / CLI Reference

FormulaBench exposes a CLI rather than an HTTP API. Display the full local reference with:

```sh
uv run python -m formulabench.capture --help
```

| Argument | Behaviour |
| --- | --- |
| `--dataset-dir PATH` | Required SpreadsheetBench dataset directory. |
| `--out-dir PATH` | Required empty output directory, or an existing run with `--resume`. |
| `--ids ID[,ID...]` | Restricts execution to validated task IDs. |
| `--concurrency N` | Maximum concurrent remote sampling calls from 1 to 64, with a default of 4. |
| `--preflight-only` | Validates the fixed native runtime without sampling the model. |
| `--resume` | Validates and continues an existing output directory. |
| `--retry-failures` | With `--resume`, reschedules failed checkpoints and may make new provider calls. |
| `--replay-write-failures` | With `--resume`, reuses eligible stored responses with zero model calls. |

There is no declared stable Python library API. Internal modules and formats may change until the
project adopts a versioned public interface.

## Tests

Run the credential-free local checks:

```sh
uv run pytest
uv run ruff check formulabench tests experiments training baseline/tinker_predict.py
uv run ruff format --check formulabench tests experiments training baseline/tinker_predict.py
uv run python evaluate.py --oracle
```

Build the same container targets used by CI:

```sh
docker build --target contract-test -t formulabench:contract-test .
docker build --target runtime -t formulabench:latest .
```

GitHub Actions also downloads the checksum-verified public dataset and verifies the all-400 target
contracts and frozen split. See the [current workflow](https://github.com/MasteraSnackin/FormulaBench/actions/workflows/ci.yml).

## Roadmap

These are proposed improvements, not claims about the frozen 400-task result:

- Route cell-level formula requests and sheet-level transformations through separate strategies.
- Compile large filtering, sorting, appending, grouping and clearing tasks into deterministic
  operations instead of model-enumerated cells.
- Make workbook retrieval task-aware and token-aware to reduce context truncation.
- Recalculate candidates in an isolated LibreOffice process and reject runtime formula errors
  before publication.
- Add formula compatibility normalisation and output invariants.
- Generate and verify the hosted evidence metrics from the hashed evaluation result during
  deployment.
- Keep the LoRA checkpoint and inference configuration frozen, then make a broader comparison on
  the historically named `held_out` and `final_only` reporting buckets.

See [`experiments/README.md`](experiments/README.md) for the failure analysis behind the retrieval,
task-routing, deterministic-operation and controlled-comparison priorities.

## Contributing

Use [GitHub Issues](https://github.com/MasteraSnackin/FormulaBench/issues) to report a defect, propose
an experiment or discuss a change. GitHub sign-in is required. The repository does not yet declare
code-contribution terms, so agree the scope with the maintainer before preparing a pull request.
Do not copy, modify or redistribute code unless the relevant rights holders have granted permission.

Never commit API keys, private workbooks, the downloaded dataset or unredacted provider data.
Any authorised change must preserve fail-closed writes, deterministic artefacts and an honest
separation between development evidence and unseen private evaluation.

## License

This repository does not currently declare a project code licence and contains no `LICENSE` file.
Do not assume permission to reuse, modify or redistribute the project code. The public dataset is
separately identified by the organisers as CC BY-SA 4.0, and the referenced Qwen model declares
Apache-2.0. Neither grants a licence for FormulaBench itself. See [PROVENANCE.md](PROVENANCE.md) for
the recorded upstream, dataset and model boundaries.

## Contact / Support

- Repository owner and declared team member: [MasteraSnackin](https://github.com/MasteraSnackin)
- Bug reports and research proposals: [GitHub Issues](https://github.com/MasteraSnackin/FormulaBench/issues)
- Repository: [github.com/MasteraSnackin/FormulaBench](https://github.com/MasteraSnackin/FormulaBench)

No public support email or service-level commitment is declared.
