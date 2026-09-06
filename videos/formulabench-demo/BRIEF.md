---
workflow: product-launch-video
flow: automation
storyboard: no
message: "FormulaBench v2 makes spreadsheet generation auditable by constraining the model, preserving failed inputs and separating migration parity from benchmark accuracy."
destination: youtube-embed
aspect: 1920x1080
language: en-GB
audience: "Encode hackathon judges, spreadsheet researchers, and technically curious fund-operations teams"
length: 75s
angle: "The audit trail"
narration: minimal
---

## Intent

A concise research-demo film for FormulaBench v2, the Research: Excel Formula Generation submission. Open on a workbook that looks healthy but is still wrong, then follow the v2 path from bounded context through a typed Qwen plan, controlled execution, fallback and organiser evaluation. The film should make the migration inspectable without presenting an unscored candidate as a benchmark result.

## Customizations

- Use the workbook as the visual world: formula bar, grid, named target cells, execution receipts and dated run state.
- Show the method as bounded context, Qwen3.8-27B typed plan, typed operations or screened Python, validation, pristine fallback, save/reopen and organiser evaluation.
- State the logical model-call policy precisely: one initial call and at most one bounded repair or truncation recovery.
- Include the credential-free migration replay: 400 outcomes checked, 369 accepted plans replayed, 31 pristine fallbacks verified and zero wrapper mismatches.
- Label that replay as migration parity rather than formula accuracy.
- Close on three separate records: FormulaBench v1 at 133/400, the historical ExactSource source run at 302/400, and FormulaBench v2 with no score claimed in this film.

## Notes

- Keep the language honest. Do not claim that FormulaBench v2 achieved ExactSource's historical 75.50% score or improved on FormulaBench v1.
- Do not treat the two selected route canaries as a representative benchmark sample.
- Do not describe screened Python as a general security boundary or claim that it enforces answer-range containment.
- The active full run began at 10:46 BST on 6 September 2026 and had not been evaluated when this brief was updated. The film deliberately makes no v2 score claim, so it remains accurate after the run ends.
- Avoid generic chatbot screens, AI sparkle motifs, stock photography, faux testimonials, inflated claims, and card-grid presentation.
- Prefer a restrained editorial palette: deep navy, warm paper, audit green, and failure red.
- The user asked to rebuild and publish the v2 artefacts, so this remains an autonomous build with a final verified render.
