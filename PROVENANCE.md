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

A later, separately labelled `experiments/public_benchmark_audit.json` opened all 400 public
golden workbooks to measure aggregate benchmark structure and the untouched-input baseline. Its
reproduction script emits no task IDs, instructions, cell coordinates or cell values. Aggregate
findings informed the general response contract, so the public 80/80/240 lanes are reporting
partitions rather than statistically untouched holdouts. The original input-only audit and frozen
split artefacts were not rewritten; only the organiser's private evaluation remains genuinely
unseen.

The native inference path uses the Qwen3.8 tokenizer and chat template for
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
