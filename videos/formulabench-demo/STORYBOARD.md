---
format: 1920x1080
duration: 75s
message: "FormulaBench v2 makes spreadsheet generation auditable by constraining the model, preserving failed inputs and separating migration parity from benchmark accuracy."
arc: "Wrong result, bounded method, controlled execution, evaluator, migration evidence, honest score record"
audience: "Encode hackathon judges, spreadsheet researchers, and fund-operations teams"
mode: autonomous
music: none
---

## Video direction

- Palette system: warm cream is the paper ground; ink is the editorial voice; warm navy is reserved for workbook and code surfaces; audit green is the one voltage moment in each frame; failure red appears only for rejected work.
- Type system: EB Garamond carries sentence-case claims and hero figures, Inter carries explanations, and JetBrains Mono carries formulae, paths, cells and receipts.
- Motion grammar: smooth long-tail settles with each element revealed on its spoken cue. Working states are finite and stop when their receipt lands. Holds stay still; no breathing cards or late camera drift.
- Rhythm: Frame 1 creates tension. Frames 2 to 6 explain the method. Frame 7 is the evidence peak. Frame 8 slows down for the qualification and repository hold.
- Composition: editorial asymmetry, visible spreadsheet structure, at least two focal points in produced frames, and all essential copy above the bottom 17% caption band.
- Never show: generic chatbot chrome, AI sparkle motifs, stock photography, purple-blue gradients, faux testimonials, unverified performance claims, front-loaded slideshow motion or independent screensaver drift.

## Frame 1 - Open does not mean correct

- scene: A workbook opens without an error, calculates a polished value and then fails its target-cell check.
- voiceover: "A workbook can open cleanly and still return the wrong investment answer."
- duration: 5s
- poster: 4.4s
- transition_in: cut
- status: animated
- src: compositions/frames/01-open-not-correct.html
- type: hook
- persuasion: Pain validation
- beat: tension
- blueprint: typewriter-reveal (Adapt)
- asset_candidates:
- focal: none; native formula bar and workbook grid
- roles: formula bar = focal; workbook grid = supporting; failed target cell = voltage
- sfx:

Adapt: keep the live type-on and in-place correction engine. Resolve on a failed target-cell receipt rather than a brand reveal.

Scene 1 (0.0-1.8s): only a warm-navy formula bar sits over an oversized worksheet grid. A valid-looking formula types on with a square caret (`discrete-text-sequence`, `context-sensitive-cursor`) in the upper third.
Scene 2 (1.8-3.6s): the result cell resolves to `£4,200,000`; an `OPENED` receipt appears beside it, then the expected value `£4,020,000` reveals on the next spoken cue. The layout opens into an asymmetric 70/30 split.
Scene 3 (3.6-5.0s): the target cell receives a restrained red `WRONG VALUE` stamp (`spring-pop-entrance`). The caret stops and the contradiction holds completely still.

narrativeRole: Establish that file validity and formula plausibility do not establish spreadsheet correctness.
keyMessage: A cleanly opened workbook can still return the wrong answer.

## Frame 2 - Bounded workbook context

- scene: The full workbook recedes while the exact sheets, ranges and task instructions selected for the model stay visible.
- voiceover: "FormulaBench v2 reads bounded workbook context, then asks Qwen three point eight, twenty-seven B for a typed plan."
- duration: 8.2s
- poster: 6.3s
- transition_in: zoom-through
- status: animated
- src: compositions/frames/02-bounded-context.html
- type: product_intro
- persuasion: Risk reduction
- beat: control
- blueprint: spatial-pan-stations (Adapt)
- asset_candidates:
- focal: none; native workbook-context strip
- roles: selected sheets = focal; excluded workbook regions = supporting; model label = voltage
- sfx:

Adapt: keep one virtual camera crossing named stations. Replace historical milestones with the exact context layers handed to the model.

Scene 1 (0.0-2.2s): the camera opens on station `01 TASK`, showing the analyst request and target range only. A hairline audit path draws towards the next station (`svg-path-draw`).
Scene 2 (2.2-4.7s): one lateral pan (`viewport-change`) lands on `02 WORKBOOK CONTEXT`; the named sheets and nearby formulas reveal sequentially while unrelated regions dim outside the bounded window.
Scene 3 (4.7-7.0s): the final pan lands on `03 QWEN3.8-27B`. A compact `TYPED PLAN REQUESTED` receipt seats beneath the model name and the camera holds.

narrativeRole: Introduce v2 as a bounded, inspectable model call rather than an opaque workbook upload.
keyMessage: The model receives task-specific workbook context and must return a typed plan.

## Frame 3 - Two execution routes

- scene: One audit rail divides into typed workbook operations for target cells and screened Python for larger sheet changes.
- voiceover: "Target cells use fixed workbook operations. Larger sheet changes may use screened Python with limited imports and resources."
- duration: 8.5s
- poster: 8.2s
- transition_in: push-slide LEFT
- status: animated
- src: compositions/frames/03-two-routes.html
- type: feature_showcase
- persuasion: Method transparency
- beat: clarity
- blueprint: compose
- asset_candidates:
- focal: none; native split execution diagram
- roles: typed operations = left focal; screened Python = right focal; shared validator = voltage
- sfx:

Scene 1 (0.0-2.7s): a single audit rail enters from the left and stops under `TYPED PLAN`; the first route label appears only as the narration reaches target cells (`dynamic-content-sequencing`).
Scene 2 (2.7-5.6s): the rail branches into a left lane labelled `TYPED WORKBOOK OPERATIONS`; `set value`, `set formula` and `copy style` reveal one at a time in a dense but flat 60/40 composition.
Scene 3 (5.6-8.0s): the right lane reveals `SCREENED PYTHON`, followed by `limited imports` and `resource caps`. The branches remain equal in height without card tilt or decorative glow.
Scene 4 (8.0-9.0s): both lanes terminate at one green `VALIDATE OUTPUT` node. The finished diagram holds still.

narrativeRole: Explain the v2 execution split without overstating the Python screen as a security sandbox.
keyMessage: V2 uses fixed operations where possible and a constrained alternative for larger transformations.

## Frame 4 - Pristine fallback

- scene: A rejected plan is stopped before write, and the original workbook passes through unchanged with an explicit fallback reason.
- voiceover: "Rejected plans never touch the submission. V2 preserves the pristine workbook and records the fallback."
- duration: 7s
- poster: 6.2s
- transition_in: squeeze
- status: animated
- src: compositions/frames/04-pristine-fallback.html
- type: feature_showcase
- persuasion: Risk reversal
- beat: trust
- blueprint: agent-progress-theater (Adapt)
- asset_candidates:
- focal: none; native validation ledger and pristine workbook receipt
- roles: rejected plan = focal; pristine workbook = supporting; fallback reason = voltage
- sfx:

Adapt: keep the working-state-to-receipt mutation. Replace a celebratory checklist with a visible rejection and untouched-workbook handoff.

Scene 1 (0.0-2.0s): a validation ledger types `CHECKING PLAN` beside a finite spinner (`discrete-text-sequence`, `svg-icon-enrichment`). Rows for target ownership, paths and limits arrive on their spoken cues.
Scene 2 (2.0-4.2s): one row turns failure red and the spinner stops immediately. A `REJECTED BEFORE WRITE` strip replaces the working label (`scale-swap-transition`).
Scene 3 (4.2-6.0s): the rejected plan collapses while a pristine workbook receipt expands into the same centre: `INPUT PRESERVED` and `FALLBACK RECORDED`.
Scene 4 (6.0-7.0s): the receipt holds with its reason visible in monospace. No hidden partial workbook appears.

narrativeRole: Show that rejected output becomes an explicit pristine fallback rather than a silent partial submission.
keyMessage: Rejected plans cannot modify the submitted workbook.

## Frame 5 - One call, one bounded second chance

- scene: A model-call counter advances once, exposes one optional repair branch and then closes before save and reopen checks.
- voiceover: "Each task gets one model call, plus at most one bounded repair, before save and reopen checks."
- duration: 7.5s
- poster: 7.2s
- transition_in: crossfade
- status: animated
- src: compositions/frames/05-bounded-calls.html
- type: feature_showcase
- persuasion: Scope control
- beat: scrutiny
- blueprint: agent-progress-theater (Adapt)
- asset_candidates:
- focal: none; native call ledger
- roles: initial call = focal; optional repair = supporting; closed call budget = voltage
- sfx:

Adapt: keep the stateful progress ledger, but make the call ceiling and save checks the receipt.

Scene 1 (0.0-2.4s): `LOGICAL CALL 1` seats beside a Qwen request strip. A single counter ticks from zero to one and stops (`counting-dynamic-scale`, fixed-size register).
Scene 2 (2.4-4.8s): a narrow branch labelled `OPTIONAL` expands only as the narration says “at most one”; it carries `semantic repair` and `truncation recovery`, never both.
Scene 3 (4.8-6.7s): the branch closes into `CALL BUDGET CLOSED`. `SAVE` and `REOPEN` checks land in sequence with drawn checkmarks (`svg-path-draw`).
Scene 4 (6.7-8.0s): the full ledger holds on the exact ceiling and completed workbook checks.

narrativeRole: Make the bounded logical-call policy and post-write verification visible.
keyMessage: V2 allows one initial call and at most one narrowly defined second call.

## Frame 6 - Evaluation at the target cells

- scene: LibreOffice recalculates the workbook, then the organiser evaluator checks the required cells against their expected values.
- voiceover: "The organiser evaluator recalculates in LibreOffice and compares every target cell. Executable can still mean wrong."
- duration: 8s
- poster: 7.3s
- transition_in: push-slide LEFT
- status: animated
- src: compositions/frames/06-target-cell-evaluation.html
- type: feature_showcase
- persuasion: Show-don't-tell proof
- beat: scrutiny
- blueprint: grid-card-assemble (Adapt)
- asset_candidates:
- focal: none; native worksheet target-cell field
- roles: target cells = focal; LibreOffice receipt = supporting; mismatch = voltage
- sfx:

Adapt: keep the sequential field population but use one worksheet surface rather than a dashboard card grid.

Scene 1 (0.0-2.4s): a workbook receipt appears at the top left: `OPEN`, `RECALCULATE`, `SAVE`. The three states reveal in order and stop on `LIBREOFFICE 26.8.0.3`.
Scene 2 (2.4-5.2s): target cells populate across a full-width worksheet field one by one (`center-outward-expansion`, short-path form). Each cell pairs `expected` and `produced` values.
Scene 3 (5.2-7.0s): most cells settle green while one exact mismatch turns red only on the final sentence. The heading changes in place to `EXECUTABLE, STILL WRONG` (`discrete-text-sequence`).
Scene 4 (7.0-8.0s): the evaluator receipt and mismatched target cell hold together for comparison.

narrativeRole: Demonstrate why workbook execution and target-cell correctness require separate evidence.
keyMessage: The evaluator judges recalculated target values, not whether the file merely runs.

## Frame 7 - Migration parity receipt

- scene: A 400-outcome replay resolves into 369 accepted plans, 31 pristine fallbacks and zero wrapper mismatches.
- voiceover: "A credential-free migration replay checked four hundred outcomes: three hundred sixty-nine accepted plans, thirty-one pristine fallbacks, and zero wrapper mismatches."
- duration: 11.5s
- poster: 10.2s
- transition_in: zoom-through
- status: animated
- src: compositions/frames/07-migration-parity.html
- type: social_proof
- persuasion: Reproducibility evidence
- beat: confidence
- blueprint: dataviz-countup (Adapt)
- asset_candidates:
- focal: none; native migration audit instrument
- roles: 400 checked = focal; 369 and 31 partition = supporting; zero mismatches = voltage
- sfx:

Adapt: keep a single count-up instrument as the hero. Replace marketing metrics with a partitioned audit receipt and a permanent scope qualification.

Scene 1 (0.0-3.3s): the kicker `CREDENTIAL-FREE MIGRATION REPLAY` appears before a large `400` count-up and thin progress ring (`counting-dynamic-scale`, `stat-bars-and-fills`).
Scene 2 (3.3-6.7s): the total resolves into two exact bands: `369 ACCEPTED PLANS REPLAYED` and `31 PRISTINE FALLBACKS VERIFIED`. The fills arrive on their spoken cues and sum visually to 400.
Scene 3 (6.7-9.4s): a large `0` settles above `WRAPPER MISMATCHES`; the prior bands remain visible but subordinate.
Scene 4 (9.4-11.0s): `MIGRATION PARITY, NOT BENCHMARK ACCURACY` reveals under a hairline rule and the complete receipt holds still.

narrativeRole: Present deterministic migration evidence while keeping its scope distinct from formula accuracy.
keyMessage: The v2 wrapper reproduced every retained ExactSource outcome in the migration replay.

## Frame 8 - Three separate score records

- scene: Three dated records appear in sequence, then clear to a FormulaBench v2 repository card.
- voiceover: "That checks migration parity, not benchmark accuracy. FormulaBench v1 passed one hundred thirty-three of four hundred. The historical ExactSource source run passed three hundred two. No FormulaBench v2 score is claimed in this film."
- duration: 19.3s
- poster: 19.0s
- transition_in: blur-crossfade 0.4s
- status: animated
- src: compositions/frames/08-score-record.html
- type: cta
- persuasion: Honest limitation
- beat: resolve
- blueprint: titlecard-reveal (Adapt)
- asset_candidates:
- focal: none; native three-record score ledger and repository lockup
- roles: score records = supporting; v2 status = focal; repository address = voltage
- sfx:

Adapt: keep the calm card-chain shape. Give each evidence state its own full-opacity record, then end on a static repository lockup.

Scene 1 (0.0-3.8s): a centred qualification holds: `MIGRATION PARITY` above `NOT BENCHMARK ACCURACY`. One restrained rule draws underneath; nothing else appears.
Scene 2 (3.8-8.0s): a hard cut at full opacity replaces it with `FORMULABENCH V1` and the record `133 / 400`, followed by `33.25% TASK PASS RATE` as the narration names it.
Scene 3 (8.0-12.5s): another hard cut reveals `HISTORICAL EXACTSOURCE SOURCE RUN`, `302 / 400` and `75.50% TASK PASS RATE`. A permanent line reads `NOT A FORMULABENCH V2 RESULT`.
Scene 4 (12.5-16.2s): the third record reads `FORMULABENCH V2`, `NO SCORE CLAIMED IN THIS FILM`, and `NEW RESULT REQUIRES HASH-BOUND EVALUATION`.
Scene 5 (16.2-20.0s): the ledger clears to the FormulaBench v2 lockup, `AUDITABLE EXCEL FORMULA GENERATION`, and `github.com/MasteraSnackin/FormulaBench`. The green spike seats once, then the card stays static to the final frame.

narrativeRole: Prevent score substitution and leave judges with a precise repository path.
keyMessage: Historical evidence is public, but FormulaBench v2 receives no score until its new run is evaluated.
