# Provenance

FormulaBench started from the Research Track scaffold published by Ylookup for the
Encode Rebuild Private Markets Hackathon.

- Upstream repository: https://github.com/ylookup/encode-hackathon
- Imported upstream commit: `37d9016264762a25cae49e077cd0893055bd9093`
- Imported `research/` tree: `cf4be017f5b099ff40142e29967b16d8f6460fe3`
- Import time: after the official 12:00 BST start on 5 September 2026
- Imported paths: every Git-tracked file below upstream `research/`

Every imported file was checked against its upstream Git blob before FormulaBench changes
began. The upstream repository did not declare a code licence at import time, so this project
does not invent or imply one. The separately downloaded
[SpreadsheetBench Verified](https://huggingface.co/datasets/KAKA22/SpreadsheetBench) dataset is
identified by the organisers as
[CC-BY-SA-4.0](https://creativecommons.org/licenses/by-sa/4.0/) and is not committed to this
repository. `experiments/public_input_audit.json` is a derived statistical summary of the
manifest and initial workbooks; its reproduction script does not open golden workbooks.

## ExactSource-backed v2

FormulaBench v2 vendors the Python source from ExactSource `src/exactsource/` as the top-level
`exactsource` package. The vendored source is current at ExactSource commit
`99fe8084bf35a5fca6a2c2e1c9beae802766a618`. The inference core is byte-identical to the core in
the ExactSource commit used for its scored run,
`8b84dba1d9263e2123b8f15267239b70ff817907`; later upstream changes did not alter those core
files. The current FormulaBench Docker base is the same pinned Python 3.12.11 image digest used by
that scored ExactSource commit. [`NOTICE-EXACTSOURCE.md`](NOTICE-EXACTSOURCE.md) records the
vendored component and its MIT licence; the notice is included in both the wheel and runtime image.

ExactSource's immutable aggregate record for that source run is
[`experiments/full_400_8b84dba.json`](https://github.com/MasteraSnackin/ExactSource/blob/99fe8084bf35a5fca6a2c2e1c9beae802766a618/experiments/full_400_8b84dba.json),
SHA-256 `81097eb0d649727b469b3d5f932464889e4bd91f88d57c4053ff446684e8c5a7`. It reports
302/400 passed tasks, a 75.50% pass rate and 80.06% cell accuracy. These are upstream
ExactSource results and are recorded only as the migration baseline; FormulaBench v2 remains
unscored.

The canonical container excludes the optional `lxml` and Pillow packages, matching the scored
ExactSource dependency environment. The v2 adapter explicitly selects openpyxl's standard-library
XML serialiser in both the parent process and ExactSource's allow-listed transformation-child
environment. FormulaBench is installed non-editably in the image so ExactSource's isolated Python
worker can import the same packaged code without depending on the working directory. A direct v2
run also checks this dependency surface before loading tasks, creating output or contacting Tinker;
it directs the operator to Docker or an exact `uv sync --locked` base environment when the check
fails.

`scripts/verify_v2_migration.sh` performs a network-disabled, credential-free reconstruction
against a separately supplied ExactSource run. On 6 September 2026 it checked all 400 retained
outcomes: 369 accepted plans, 31 byte-identical fallbacks and zero workbook-content mismatches. It
compares exact OOXML member-name inventories and decompressed member payloads, normalising only the
core creation/modification timestamp values generated at save time. ZIP container ordering and
metadata are intentionally outside the comparison; worksheet, formula, relationship, drawing,
style and all other OOXML payload content remains exact. It neither opens golden workbooks nor calls
a provider. The result proves deterministic migration parity only.
`experiments/v2_migration_validation.json` records its method and the two separate paid route
canaries without presenting either as a full benchmark score.

This lineage is a source and runtime-parity statement only. FormulaBench v2 has not yet completed
and been evaluated across all 400 tasks, so it has no published FormulaBench score and makes no
claim of matching or improving on either FormulaBench v1 or ExactSource. The existing 33.25%
result and all committed score artefacts remain historical v1 evidence. A complete 400-task
evaluator result is required before any v2 performance claim.

A later, separately labelled `experiments/public_benchmark_audit.json` opened all 400 public
golden workbooks to measure aggregate benchmark structure and the untouched-input baseline. Its
reproduction script emits no task IDs, instructions, cell coordinates or cell values. Aggregate
findings informed the general response contract, so the public 80/80/240 lanes are reporting
partitions rather than statistically untouched holdouts. The original input-only audit and frozen
split artefacts were not rewritten; only the organiser's private evaluation remains genuinely
unseen.

The historical v1 native inference path uses the Qwen3.8 tokenizer and chat template for
`Qwen/Qwen3.8-27B` at immutable Hugging Face revision
`1d4bf0f2ff6012fd82039f2fa52739d0dd7c60c0`. The model repository identifies its
licence as Apache-2.0. FormulaBench does not redistribute model weights. The submitted Docker
build downloads only the four tokenizer/template files at that pinned revision, verifies these
SHA-256 values and retains them in the image for offline runtime loading:

- `chat_template.jinja`: `c3cf9e34abf4f9e36c2d72165aa9c132d3e2a725b6c2586aaa3a8af9d7a81041`
- `config.json`: `191e0af232104ed8b65258cf3fb2b842e288008baca7633c11b82a1ac7203aab`
- `tokenizer.json`: `0997f410c57a1f4e53b09e4be8f4a172d90edd9564368fb0847030937229b9f3`
- `tokenizer_config.json`: `b11349aafa7cdc6a320767cf7ceb29ed82f7eda5d65e8e0819e76f0ce947bf27`

The production dependency wheels for Tinker and Transformers include their Apache-2.0 licence
files in the installed distributions. The optional `tinker-cookbook` comparison environment is
not installed in the submitted runtime image.
