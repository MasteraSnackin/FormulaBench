---
format: 1920x1080
duration: 75.050s
message: "FormulaBench turns spreadsheet-generation claims into reproducible, cell-level evidence, then separates the frozen public score from a later Tinker development experiment."
arc: "Hook → Standard of proof → Method → Safety gate → Public results → Failure insight → Post-score Tinker experiment → Next step"
audience: "Encode hackathon judges, spreadsheet researchers, and fund-operations teams"
mode: autonomous
music: none
---

## Video direction

- Palette system: warm cream is the default paper ground; ink is the editorial voice; warm navy is reserved for formulas and workbook surfaces; audit green is the single voltage moment in each frame; failure red appears only as semantic status inside the navy surface.
- Type system: EB Garamond carries sentence-case claims and hero figures, Inter carries explanations, and JetBrains Mono carries formulae, cells, stage labels, and evaluator receipts.
- Motion grammar: smooth long-tail settles, one purposeful move at a time, with every reveal paced to the spoken cue across the back half of the shot. Working states are finite and die when their receipt lands. Holds are genuinely still; no breathing cards or late camera drift.
- Rhythm: controlled–sharp–mechanical–measured–peak–hold–qualified–resolve. Frames 2 and 6 are deliberate held reads; Frame 5 is the public-score peak; Frame 7 is a compact post-score checkpoint; Frame 8 gives the URL the longest hold.
- Composition: editorial asymmetry, visible grid structure, two focal points per produced frame, and all essential copy above the bottom 17% caption band.
- Never show: generic chatbot chrome, floating glass cards, AI sparkle motifs, stock photography, purple-blue gradients, fake testimonials, unverified baseline comparisons, front-loaded slideshow motion, or independent screensaver drift.

## Frame 1 - Plausible is not correct

- scene: An Excel-style formula types into a warm-navy formula bar, computes a polished but wrong result, and receives an emphatic REJECTED audit stamp.
- voiceover: "A formula can look right, reference real cells, and still return the wrong answer."
- duration: 4.702s
- poster: 4.2s
- transition_in: cut
- status: animated
- src: compositions/frames/01-plausible-not-correct.html
- type: hook
- persuasion: Pain validation
- beat: tension
- blueprint: typewriter-reveal (Adapt)
- asset_candidates:
- focal: none; native formula-bar and worksheet composition
- roles: native formula bar = focal; worksheet cells = supporting; audit stamp = voltage
- sfx:

Adapt: keep the live type-on and in-place correction engine; replace the brand-pop resolve with a semantic audit stamp, because the story is about proof rather than a product reveal.

Scene 1 (0.0–1.6s): only the formula bar is visible over an oversized worksheet grid; `=IFERROR(INDEX(Returns!$F:$F,MATCH(A12,Returns!$A:$A,0)),0)` types on character by character with a square caret (`discrete-text-sequence`, `context-sensitive-cursor`) in a rule-of-thirds upper band, with the result cell still blank below.
Scene 2 (1.6–3.4s): as the narration says “reference real cells”, the result cell resolves to `£4,200,000`; a small green syntax chip reads `VALID`, while a red semantic note reveals one cue later: `Expected £4,020,000` (`discrete-text-sequence`). The worksheet expands into an asymmetric 70/30 composition with formula left and evidence right.
Scene 3 (3.4–4.702s): on “wrong answer”, `REJECTED` stamps across the result in one smooth scale settle (`spring-pop-entrance`, restrained register), the caret dies, and the frame holds completely still on the contradiction.

narrativeRole: Establish that plausible syntax and real references are not enough; the benchmark must test the answer.
keyMessage: A formula that looks right can still be wrong.

## Frame 2 - Raise the standard

- scene: “Valid syntax” is replaced in place by “Correct target cells”, with FormulaBench named as the evidence layer beneath.
- voiceover: "That’s why FormulaBench doesn’t stop at valid syntax. It asks whether every target cell is actually correct."
- duration: 6.844s
- poster: 6.1s
- transition_in: zoom-through
- status: animated
- src: compositions/frames/02-raise-the-standard.html
- type: benefit_highlight
- persuasion: Negative contrast
- beat: clarity
- blueprint: kinetic-type-beats (Adapt)
- asset_candidates:
- focal: none; native kinetic typography and cell coverage rail
- roles: fixed statement = focal; coverage rail = supporting; FormulaBench wordmark = voltage
- sfx:

Adapt: keep the fixed-line token swap signature, but use a measured editorial cadence rather than a rapid product-name flash.

Scene 1 (0.0–2.3s): a fixed serif line lands left of centre: `A workbook can have` while the mono token `VALID SYNTAX` seats beside it via a per-word reveal (`dynamic-content-sequencing`). A thin 1px rule draws beneath the line.
Scene 2 (2.3–5.3s): on “doesn’t stop”, `VALID SYNTAX` cuts out in place and `CORRECT TARGET CELLS` replaces it at the same anchor (`discrete-text-sequence`); beneath, cells A12, B12, C12, and D12 reveal sequentially and lock green only on the words “every target cell”.
Scene 3 (5.3–6.844s): the label `FORMULABENCH / RESEARCH: EXCEL FORMULA GENERATION` appears in mono at the top edge; the main statement and cell rail hold still for the final read.

narrativeRole: State the value claim by beat two: FormulaBench measures correctness at the cells that matter.
keyMessage: The standard is exact target-cell correctness, not syntactic validity.

## Frame 3 - One audited path

- scene: A single worksheet-width audit line moves through six named stages from workbook inspection towards the organiser-supplied evaluator.
- voiceover: "The harness inspects the workbook, builds sheet-aware context, samples Qwen once, then validates exact coverage and formula safety."
- duration: 8.594s
- poster: 7.8s
- transition_in: push-slide LEFT
- status: animated
- src: compositions/frames/03-audited-path.html
- type: product_intro
- persuasion: Show-don’t-tell proof
- beat: control
- blueprint: spatial-pan-stations (Adapt)
- asset_candidates:
- focal: none; native workbook audit stations
- roles: current station = focal; audit line = supporting; stage receipt = voltage
- sfx:

Adapt: keep the single virtual traversal and terminal held station; compress six stages into three paired stops so the path remains legible at video scale.

Scene 1 (0.0–2.6s): an oversized worksheet world opens on paired stations `01 INSPECT` and `02 CONTEXT`; the audit line draws between them and the world pans left to centre each cue (`viewport-change`, `svg-path-draw`) while sheet tabs and named ranges assemble as supporting detail.
Scene 2 (2.6–5.8s): the camera continues on the same axis to `03 SAMPLE ONCE`; a single Qwen request strip locks in, then `04 VALIDATE COVERAGE` reveals only when named, with a four-cell target set snapping into exact alignment (`coordinate-target-zoom`, `dynamic-content-sequencing`).
Scene 3 (5.8–8.594s): the final pan lands on `05 FORMULA SAFETY`; sheet-qualified references and the external-reference gate tick on in sequence, then the line terminates at a green `CONTRACT VALID` receipt and holds (`stat-bars-and-fills`, no camera motion after 7.8s).

narrativeRole: Make the method inspectable and reproducible rather than presenting an unexplained model answer.
keyMessage: FormulaBench uses one controlled, sheet-aware, contract-checked generation path.

## Frame 4 - Write, or fail closed

- scene: A validation ledger checks five constraints, then branches cleanly to an atomic workbook write or a visible fallback before evaluation.
- voiceover: "Only then does it write atomically, or fail closed, before the organiser-supplied evaluator scores the workbook."
- duration: 7.393s
- poster: 5.6s
- transition_in: squeeze
- status: animated
- src: compositions/frames/04-write-or-fail-closed.html
- type: feature_showcase
- persuasion: Risk reversal
- beat: trust
- blueprint: agent-progress-theater (Adapt)
- asset_candidates:
- focal: none; native validation ledger
- roles: ledger = focal; write/fallback branch = supporting; evaluator receipt = voltage
- sfx:

Adapt: keep the working-state-to-receipt mutation, but remove the cursor and decorative loader; the validator itself is the actor.

Scene 1 (0.0–1.6s): a warm-navy ledger enters with one active row only: `Target coverage`; its indicator draws to completion, then `Sheet exists` and `Formula allowed` cascade in as the narration reaches “only then” (`svg-path-draw`, `dynamic-content-sequencing`).
Scene 2 (1.6–4.0s): five rows mutate from numbered outline to green checks one by one; a pending `Reopen and verify` row finishes last, and the working indicator dies immediately (`scale-swap-transition`, `svg-path-draw`).
Scene 3 (4.0–5.0s): the ledger splits into two equal paths: `ATOMIC REPLACE` in audit green and `FAIL CLOSED` in ink; both are explicit, neither is hidden. A small `NO SILENT PARTIALS` label seats between them.
Scene 4 (5.0–7.393s): the paths collapse into one receipt reading `WORKBOOK READY FOR SELF-EVALUATION`; it holds with no idle motion. The footer records the actual order: temporary save, reopen and verify, atomic replace, then evaluator.

narrativeRole: Demonstrate that invalid or partial output cannot silently masquerade as success.
keyMessage: The harness either writes a re-opened workbook atomically or fails closed before grading.

## Frame 5 - Public self-evaluation

- scene: Our public self-evaluation of 400 benchmark tasks builds as one large pass-rate instrument, followed by cell accuracy and coverage receipts.
- voiceover: "In our public self-evaluation of four hundred benchmark tasks, one hundred thirty-three passed: a thirty-three point two five percent pass rate, with forty point five six percent cell accuracy."
- duration: 12.722s
- poster: 10.4s
- transition_in: zoom-through
- status: animated
- src: compositions/frames/05-public-self-evaluation.html
- type: social_proof
- persuasion: Statistical proof
- beat: confidence
- blueprint: dataviz-countup (Adapt)
- asset_candidates:
- focal: none; native evaluator instrument
- roles: pass rate = focal; passed tasks and cell accuracy = supporting; public self-evaluation tag = voltage
- sfx:

Adapt: keep the number-and-graphic landing as one beat, but use a single flat ring plus editorial receipts instead of a perspective dashboard.

Scene 1 (0.0–5.9s): the mono kicker `PUBLIC SELF-EVALUATION / 400 BENCHMARK TASKS` reveals first. The denominator `/ 400` stays fixed; the serif numerator and thin circular sweep begin only as the narration reaches `133`, then count from `0 → 133` (`counting-dynamic-scale`, `stat-bars-and-fills`, `svg-path-draw`).
Scene 2 (5.9–9.7s): the hero `33.25% PASS RATE` arrives on its spoken cue while a precise progress bar fills to the same value. No extra claim appears.
Scene 3 (9.7–12.2s): only when spoken, `40.56% CELL ACCURACY` draws in beneath a 1px rule; the pass figure remains dominant at a 3:1 scale ratio.
Scene 4 (12.2–12.722s): all three receipts settle into one evidence lockup and hold completely still.

narrativeRole: Put the verified public self-evaluation result at the visual peak without inflating or obscuring it.
keyMessage: The 400-task run passed 133 tasks with 40.56% cell accuracy.

## Frame 6 - Running is not right

- scene: Two apparently valid workbooks open side by side; the semantic audit reveals that most wrong cells lived inside contract-valid files.
- voiceover: "There were zero missing predictions and zero evaluator errors. But ninety-four point eight four percent of wrong cells sat inside contract-valid workbooks. Running is not the same as right."
- duration: 13.531s
- poster: 12.0s
- transition_in: blur-crossfade 0.4s
- status: animated
- src: compositions/frames/06-running-is-not-right.html
- type: feature_showcase
- persuasion: Honest limitation
- beat: scrutiny
- blueprint: comparison-split (Adapt)
- asset_candidates:
- focal: none; paired native workbook surfaces
- roles: contract-valid workbook = left focal; semantic audit = right focal; 94.84% receipt = voltage
- sfx:

Adapt: keep the equal-weight split and paired reveal; remove the 3D tilt because the design system forbids tilt, using mirrored lateral arrivals and hairline workbook frames instead.

Scene 1 (0.0–4.0s): the shared heading `COMPLETE RUN` reveals above two empty workbook frames; mono receipts `0 MISSING PREDICTIONS` and `0 EVALUATOR ERRORS` land at opposite inner edges in sync with the first sentence (`dynamic-content-sequencing`).
Scene 2 (4.0–5.2s): both illustrative workbook surfaces slide in from opposite sides on mirrored paths (`split-tilt-cards`, tilt suppressed); their top status bars both read `CONTRACT VALID`.
Scene 3 (5.2–10.9s): the right workbook's semantic cells turn failure red one by one as the serif figure `94.84%` lands between the two surfaces with the mono unit `OF WRONG CELLS` (`counting-dynamic-scale`, `discrete-text-sequence`).
Scene 4 (10.9–13.531s): the workbooks dim and the line `Running is not the same as right.` takes the centre in EB Garamond italic; the split stops moving and holds.

narrativeRole: Show the most important limitation: contract checks catch malformed output, not semantic error.
keyMessage: Most wrong cells were inside files that satisfied the output contract.

## Frame 7 - Post-score Tinker checkpoint

- scene: A dated provider capture establishes the training run, then yields to a FormulaBench-authored development A/B summary with a permanent qualification.
- voiceover: "After scoring, Tinker lifted development passes from seven to nine. This is checkpoint selection, not held-out proof."
- duration: 7.680s
- poster: 6.7s
- transition_in: blur-crossfade 0.4s
- status: animated
- src: compositions/frames/07-tinker-development-evidence.html
- type: social_proof
- persuasion: Qualified experimental evidence
- beat: qualification
- blueprint: viewport-change + theme-crossfade-morph (Adapt)
- asset_candidates:
- focal: sanitised Tinker capture followed by the controlled development comparison
- roles: provider capture = provenance; A/B result = focal; permanent caveat = trust guardrail
- sfx:

Adapt: preserve the real provider pixels inside the editorial frame, using a restrained camera move and a single crossfade rather than recreating the dashboard or animating unsupported throughput figures.

Scene 1 (0.0–3.0s): the mono qualification `POST-SCORE DEVELOPMENT CHECKPOINT SELECTION` and dated source label appear before the sanitised provider capture. The capture is fitted without cropping its axes; a subtle 1.06× camera push centres the model, rank and non-zero activity charts (`viewport-change`).
Scene 2 (3.0–6.2s): the provider capture crossfades to the FormulaBench-authored comparison at pixel-identical geometry (`theme-crossfade-morph`). `7 / 15 → 9 / 15` becomes the sole large figure, while `+2 net passes` remains subordinate.
Scene 3 (6.2–7.680s): the comparison holds with `NOT A HELD-OUT GENERALISATION RESULT` fixed above the caption band. No exact throughput or utilisation average is shown because the provider view does not supply one.

narrativeRole: Record the post-score Tinker experiment without presenting development checkpoint selection as the frozen public benchmark.
keyMessage: A rank-32 LoRA improved this controlled 15-task development comparison from seven passes to nine, but this is not held-out evidence.

## Frame 8 - Execution-aware checks

- scene: The next research experiment resolves into two concrete checks, then clears to a factual FormulaBench end card and repository address.
- voiceover: "The next step is targeted: deterministic transformations where rules are reliable, plus formula-execution checks where semantics decide the answer."
- duration: 13.584s
- poster: 10.4s
- transition_in: crossfade 0.4s
- status: animated
- src: compositions/frames/08-evidence-before-confidence.html
- type: cta
- persuasion: Future pacing
- beat: resolve
- blueprint: titlecard-reveal (Adapt)
- asset_candidates:
- focal: none; native closing lockup
- roles: two next-step lanes = supporting; FormulaBench lockup = focal; repository address = voltage
- sfx:

Adapt: keep the calm card-chain and long final hold; use one precise two-lane research card before the wordmark instead of a sales CTA.

Scene 1 (0.0–4.1s): a clean cream field holds `NEXT / RESEARCH STEP`; two horizontal lanes reveal sequentially, `DETERMINISTIC TRANSFORMATIONS` first and `FORMULA-EXECUTION VERIFICATION` second. Each has one thin progress line and no card shadow (`discrete-text-sequence`, `stat-bars-and-fills`).
Scene 2 (4.1–9.35s): the lanes slide up and crossfade into the specific experiment: `Use fixed operations for known transformations. Run generated formulas and check their results.` The meta label reads `NEXT EXPERIMENT / EXECUTION-AWARE CHECKS`.
Scene 3 (9.35–13.584s): after the spoken sentence ends, the frame cuts at full opacity to `FormulaBench` with `RESEARCH: EXCEL FORMULA GENERATION` above, `Workbook outputs and traces are public.` beneath, and `github.com/MasteraSnackin/FormulaBench` below. The audit-green spike seats once, then the lockup holds completely static to the last frame.

narrativeRole: Connect the observed semantic error rate to two testable changes in the next benchmark run.
keyMessage: The next experiment combines verified transformations with executed-formula checks.
