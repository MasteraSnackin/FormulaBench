---
workflow: product-launch-video
flow: automation
storyboard: no
message: "FormulaBench turns spreadsheet-generation claims into reproducible, cell-level evidence, then separates the frozen public score from a later development experiment."
destination: youtube-embed
aspect: 1920x1080
language: en-GB
audience: "Encode hackathon judges, spreadsheet researchers, and technically curious fund-operations teams"
length: 75s
angle: "The audit trail"
narration: minimal
---

## Intent

A concise research-demo film for FormulaBench, the Research: Excel Formula Generation submission. Open on a plausible Excel formula that is still wrong, then follow the audit trail from workbook inspection through model sampling and strict validation to the organiser-supplied evaluator. The film should build trust through evidence rather than marketing claims.

## Customizations

- Use the spreadsheet itself as the visual world: formula bar, grid, sheet-qualified references, audit stamps, and measured results.
- Show the method as inspect → context → sample → validate → atomic write/fallback → organiser-supplied evaluator.
- Include only verified figures: 133/400 passed, 33.25% pass rate, 40.56% cell accuracy, and zero missing outputs or evaluator errors.
- Include the key failure insight: 94.84% of wrong cells were in contract-valid but semantically wrong workbooks.
- After the public score and failure analysis, show the Tinker rank-32 LoRA as post-score development checkpoint-selection evidence: 7/15 to 9/15 tasks, never as a replacement public benchmark.
- Close on the next research step: deterministic transformations plus formula-execution verification.

## Notes

- Keep the language honest. Do not claim improvement over the organiser's approximately 59% baseline because the execution settings and evaluation path are not comparable.
- Keep the frozen public 400-task score visually and verbally separate from the later 15-task development comparison. The dashboard proves training activity and configuration; the FormulaBench A/B run supplies the task and cell figures.
- Avoid generic chatbot screens, AI sparkle motifs, stock photography, faux testimonials, inflated claims, and card-grid presentation.
- Prefer a restrained editorial palette: deep navy, warm paper, audit green, and failure red.
- The user asked to publish all three submission artefacts now, so this is an autonomous build with a final verified render.
