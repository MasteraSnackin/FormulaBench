# FormulaBench video source

This directory contains the reproducible HyperFrames source for the FormulaBench research demo and
the separate Tinker evidence clip. The checked-in MP4 files served by GitHub Pages are delivery
artefacts; these source compositions make their claims, timing and provenance reviewable.

## What is separated

- The main film explains the current v2 path: bounded workbook context, typed planning, two
  execution routes, pristine fallback and target-cell evaluation.
- Its 400/369/31/0 receipt is migration parity, not benchmark accuracy.
- It keeps three records separate: historical FormulaBench v1 at 133/400, the historical
  ExactSource source run at 302/400, and FormulaBench v2 with no score claimed in the film.
- The standalone Tinker evidence clip reports a separate controlled 15-task development
  comparison: 7/15 base passes and 9/15 checkpoint passes.
- The development comparison is labelled checkpoint-selection evidence, not a held-out
  generalisation result or a replacement public score.
- The provider capture establishes the model, LoRA configuration and non-zero training activity.
  FormulaBench's A/B artefact supplies the spreadsheet-correctness figures. No numerical average
  throughput or utilisation claim is made.

## Local review

Requirements: Node.js and internet access for the pinned HyperFrames runtime and GSAP dependency.
No Tinker or cloud credential is required to review or render the retained compositions.

```bash
npm run check
npm run dev
```

HyperFrames Studio opens the assembled main film. Review the claims, captions, transitions and final
repository card before rendering.

## Render commands

```bash
npm run render -- --quality high --output renders/FormulaBench-Demo.raw.mp4
npm run render -- --composition compositions/tinker-evidence.html --quality high --output renders/FormulaBench-Tinker-Evidence.raw.mp4

# Finalise the delivery files without re-encoding.
ffmpeg -i renders/FormulaBench-Demo.raw.mp4 -map 0:v:0 -map 0:a:0 -c copy -movflags +faststart renders/FormulaBench-Demo.mp4
ffmpeg -i renders/FormulaBench-Tinker-Evidence.raw.mp4 -t 30.233333 -map 0:v:0 -map 0:a:0 -c copy -movflags +faststart renders/FormulaBench-Tinker-Evidence.mp4
```

The main composition is exactly 75.00 seconds; its lossless remux adds fast-start metadata without
trimming any frame or audio sample. The standalone Tinker evidence composition is 30.25 seconds.
The Tinker finalisation step removes the renderer's terminal capture packet, so that delivery file
ends on its intended content frame while retaining the complete 30.25-second audio track.
Generated renders, snapshots, browser caches and temporary captures are intentionally ignored.
The final main-film narration WAV is 19.3 seconds long and includes the closing repository-card
hold. Main-film durations and word timings are retained in the audio metadata.

## Source map

- `BRIEF.md` defines the communication objective and evidence guardrails.
- `STORYBOARD.md` fixes scene order, duration, claims and transitions.
- `SCRIPT.md` contains the spoken copy.
- `index.html` is the assembled main composition.
- `compositions/frames/` contains the eight main-film scenes.
- `compositions/tinker-evidence.html` is the standalone evidence clip.
- `audio_meta.json` and `caption_groups.json` retain the main narration timing.
- `.media/manifest.jsonl` records the local media inventory.
- `assets/evidence/` contains only the sanitised provider capture and FormulaBench A/B summary.

The authenticated full-page Tinker screenshot and any temporary signed download URL are deliberately
excluded from this source tree.
