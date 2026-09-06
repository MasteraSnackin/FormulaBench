import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const artifactToolModule = process.env.ARTIFACT_TOOL_MODULE ?? "@oai/artifact-tool";
const { Presentation, PresentationFile } = await import(artifactToolModule);

const scriptDir = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.resolve(scriptDir, "..");
const outputPath = path.resolve(
  process.argv[2] ?? path.join(repoRoot, ".tmp-presentation", "v2", "FormulaBench-Hackathon-Deck-v2.candidate.pptx"),
);
const inspectPath = outputPath.replace(/\.pptx$/i, ".inspect.ndjson");

const W = 1280;
const H = 720;
const FONT = "Arial";
const MONO = "Menlo";
const C = {
  navy: "#0B172A",
  navy2: "#13243B",
  ink: "#122033",
  offWhite: "#F4F0E8",
  paper: "#FFFDF8",
  green: "#35C47A",
  greenDark: "#176B43",
  mint: "#DDF6E8",
  red: "#E76464",
  paleRed: "#FBE6E4",
  blue: "#4A9BC2",
  paleBlue: "#E2F1F7",
  gold: "#E4B94F",
  paleGold: "#F8EFCF",
  muted: "#667487",
  line: "#C9D1D9",
  lineDark: "#344257",
  white: "#FFFFFF",
};

const deck = Presentation.create({ slideSize: { width: W, height: H } });

function text(slide, value, position, options = {}) {
  const shape = slide.shapes.add({
    geometry: "textbox",
    name: options.name,
    position,
    fill: "none",
    line: { fill: "none", width: 0 },
  });
  shape.text = value;
  shape.text.style = {
    typeface: options.typeface ?? FONT,
    fontSize: options.fontSize ?? 22,
    bold: options.bold ?? false,
    italic: options.italic ?? false,
    color: options.color ?? C.ink,
    alignment: options.alignment ?? "left",
    verticalAlignment: options.verticalAlignment ?? "top",
    lineSpacing: options.lineSpacing ?? 1.06,
    autoFit: options.autoFit ?? "shrinkText",
    wrap: options.wrap ?? "square",
    insets: options.insets ?? { top: 0, right: 0, bottom: 0, left: 0 },
  };
  if (options.link) {
    shape.text.get(value).link = { uri: options.link, isExternal: true };
  }
  return shape;
}

function rect(slide, position, options = {}) {
  return slide.shapes.add({
    geometry: options.geometry ?? "rect",
    name: options.name,
    position,
    fill: options.fill ?? "none",
    line: options.line ?? { style: "solid", fill: C.line, width: 1 },
  });
}

function rule(slide, left, top, width, colour = C.line, thickness = 1) {
  return slide.shapes.add({
    geometry: "line",
    position: { left, top, width, height: 0 },
    fill: "none",
    line: { style: "solid", fill: colour, width: thickness },
  });
}

function title(slide, value, dark = false) {
  text(slide, value, { left: 64, top: 44, width: 1152, height: 58 }, {
    fontSize: 42,
    bold: true,
    color: dark ? C.white : C.ink,
    name: "slide-title",
  });
}

function footer(slide, number, dark = false) {
  const lineColour = dark ? C.lineDark : C.line;
  const copyColour = dark ? "#AAB6C5" : C.muted;
  rule(slide, 64, 675, 1152, lineColour, 1);
  text(slide, "FORMULABENCH V2  ·  RESEARCH: EXCEL FORMULA GENERATION", { left: 64, top: 688, width: 720, height: 18 }, {
    fontSize: 12,
    bold: true,
    color: copyColour,
    verticalAlignment: "middle",
  });
  text(slide, String(number).padStart(2, "0"), { left: 1150, top: 688, width: 66, height: 18 }, {
    fontSize: 12,
    bold: true,
    color: copyColour,
    alignment: "right",
    verticalAlignment: "middle",
  });
}

function notes(slide, lines) {
  slide.speakerNotes.textFrame.setText(lines.join("\n"));
}

function tableRuns(rows) {
  return rows.map((row, rowIndex) => row.map((value) => [[{
    run: String(value),
    textStyle: {
      typeface: FONT,
      fontSize: rowIndex === 0 ? "16px" : "17px",
      bold: rowIndex === 0,
      color: rowIndex === 0 ? C.white : C.ink,
    },
  }]]));
}

function styleTable(table, rowCount, colCount, options = {}) {
  table.borders.assign({ style: "solid", fill: C.line, width: 1 });
  for (let row = 0; row < rowCount; row += 1) {
    for (let col = 0; col < colCount; col += 1) {
      const cell = table.getCell(row, col);
      cell.fill = row === 0 ? C.navy : (options.bodyFill ?? C.paper);
      cell.text.style = {
        typeface: FONT,
        fontSize: row === 0 ? 16 : 17,
        bold: row === 0 || (options.boldFirstColumn && col === 0),
        color: row === 0 ? C.white : C.ink,
        verticalAlignment: "middle",
        lineSpacing: 1.0,
        autoFit: "shrinkText",
        insets: { top: 7, right: 9, bottom: 7, left: 9 },
      };
    }
  }
}

function metric(slide, value, label, left, top, colour = C.ink, width = 245) {
  text(slide, value, { left, top, width, height: 68 }, {
    fontSize: 52,
    bold: true,
    color: colour,
  });
  text(slide, label, { left: left + 3, top: top + 68, width, height: 34 }, {
    fontSize: 16,
    bold: true,
    color: colour === C.ink ? C.muted : colour,
  });
}

// 1. Cover
{
  const slide = deck.slides.add();
  slide.background.fill = C.navy;
  text(slide, "FORMULABENCH", { left: 66, top: 62, width: 500, height: 28 }, {
    fontSize: 15,
    bold: true,
    color: C.green,
    verticalAlignment: "middle",
  });
  text(slide, "FormulaBench v2", { left: 64, top: 148, width: 700, height: 92 }, {
    fontSize: 67,
    bold: true,
    color: C.white,
    lineSpacing: 0.94,
  });
  text(slide, "Guarded workbook execution for Qwen3.8-27B", { left: 66, top: 258, width: 670, height: 62 }, {
    fontSize: 29,
    color: "#D9E0E8",
  });
  rule(slide, 66, 350, 535, C.green, 5);
  text(slide, "Research: Excel Formula Generation\nEncode x Ylookup Rebuild Private Markets Hackathon", { left: 66, top: 383, width: 570, height: 76 }, {
    fontSize: 21,
    color: "#AAB6C5",
    lineSpacing: 1.18,
  });

  rect(slide, { left: 770, top: 133, width: 420, height: 350 }, {
    fill: C.navy2,
    line: { style: "solid", fill: C.lineDark, width: 1 },
  });
  text(slide, "TYPED PLAN", { left: 798, top: 159, width: 220, height: 25 }, {
    fontSize: 14,
    bold: true,
    color: C.green,
  });
  text(slide, "{\n  \"route\": \"operations\",\n  \"edits\": [\n    {\n      \"sheet\": \"Returns\",\n      \"cell\": \"H9\"\n    }\n  ]\n}", { left: 798, top: 199, width: 354, height: 250 }, {
    typeface: MONO,
    fontSize: 22,
    color: "#C6D2DF",
    lineSpacing: 1.11,
  });
  text(slide, "ILLUSTRATIVE STRUCTURE", { left: 798, top: 452, width: 270, height: 23 }, {
    fontSize: 12,
    bold: true,
    color: "#8F9CAE",
  });

  rect(slide, { left: 66, top: 596, width: 370, height: 42 }, {
    fill: C.green,
    line: { fill: "none", width: 0 },
  });
  text(slide, "CURRENT DEFAULT CANDIDATE", { left: 86, top: 607, width: 330, height: 20 }, {
    fontSize: 14,
    bold: true,
    color: C.navy,
    verticalAlignment: "middle",
  });
  text(slide, "UNSCORED AT CAPTURE", { left: 458, top: 607, width: 260, height: 20 }, {
    fontSize: 14,
    bold: true,
    color: C.gold,
    verticalAlignment: "middle",
  });
  text(slide, "6 September 2026", { left: 965, top: 607, width: 225, height: 20 }, {
    fontSize: 14,
    color: "#AAB6C5",
    alignment: "right",
    verticalAlignment: "middle",
  });
  footer(slide, 1, true);
  notes(slide, [
    "Sources:",
    "SUBMISSION.md section Version and evidence status",
    "formulabench/v2.py",
    "exactsource/config.py",
    "The typed plan shown here is illustrative and contains no benchmark answer.",
    "Capture status: FormulaBench v2 is the current default candidate and has no published 400-task score.",
  ]);
}

// 2. Evidence boundaries
{
  const slide = deck.slides.add();
  slide.background.fill = C.paper;
  title(slide, "Three records answer three different questions");
  text(slide, "FormulaBench v2 keeps execution evidence separate from benchmark accuracy.", { left: 64, top: 108, width: 1100, height: 38 }, {
    fontSize: 23,
    color: C.muted,
  });
  const cards = [
    {
      left: 64,
      colour: C.blue,
      eyebrow: "ENGINE",
      heading: "Can the path run?",
      body: "Typed operations and screened Python execute under bounded checks, then save to a fresh copy.",
      proof: "Code, tests and route canaries",
    },
    {
      left: 454,
      colour: C.green,
      eyebrow: "MIGRATION PARITY",
      heading: "Did the wrapper preserve behaviour?",
      body: "A credential-free replay compared retained ExactSource outcomes with the FormulaBench v2 wrapper.",
      proof: "400 checked, zero mismatches",
    },
    {
      left: 844,
      colour: C.red,
      eyebrow: "BENCHMARK",
      heading: "Are the answers correct?",
      body: "Only a complete output scored by the unchanged organiser evaluator can establish v2 accuracy.",
      proof: "Fresh v2 result not yet published",
    },
  ];
  for (const card of cards) {
    rect(slide, { left: card.left, top: 178, width: 342, height: 380 }, {
      fill: C.offWhite,
      line: { style: "solid", fill: C.line, width: 1 },
    });
    rect(slide, { left: card.left, top: 178, width: 342, height: 8 }, {
      fill: card.colour,
      line: { fill: "none", width: 0 },
    });
    text(slide, card.eyebrow, { left: card.left + 26, top: 211, width: 290, height: 23 }, {
      fontSize: 13,
      bold: true,
      color: card.colour,
    });
    text(slide, card.heading, { left: card.left + 26, top: 252, width: 290, height: 76 }, {
      fontSize: 29,
      bold: true,
      color: C.ink,
      lineSpacing: 1.02,
    });
    text(slide, card.body, { left: card.left + 26, top: 350, width: 290, height: 108 }, {
      fontSize: 19,
      color: C.muted,
      lineSpacing: 1.15,
    });
    rule(slide, card.left + 26, 485, 290, C.line, 1);
    text(slide, card.proof, { left: card.left + 26, top: 505, width: 290, height: 34 }, {
      fontSize: 16,
      bold: true,
      color: C.ink,
    });
  }
  rect(slide, { left: 64, top: 594, width: 1122, height: 54 }, {
    fill: C.navy,
    line: { fill: "none", width: 0 },
  });
  text(slide, "These records are reported separately. None is converted into a FormulaBench v2 score.", { left: 90, top: 608, width: 1070, height: 28 }, {
    fontSize: 20,
    bold: true,
    color: C.white,
    verticalAlignment: "middle",
  });
  footer(slide, 2);
  notes(slide, [
    "Sources:",
    "SUBMISSION.md sections Version and evidence status and Current v2 model path",
    "experiments/v2_migration_validation.json",
    "The migration replay and selected canaries are engineering evidence, not benchmark estimates.",
  ]);
}

// 3. Two bounded routes
{
  const slide = deck.slides.add();
  slide.background.fill = C.offWhite;
  title(slide, "One model, two bounded execution routes");
  text(slide, "Qwen/Qwen3.8-27B produces a typed plan from bounded workbook context.", { left: 64, top: 108, width: 1080, height: 38 }, {
    fontSize: 23,
    color: C.muted,
  });

  rect(slide, { left: 64, top: 174, width: 1152, height: 78 }, {
    fill: C.navy,
    line: { fill: "none", width: 0 },
  });
  text(slide, "QWEN/QWEN3.8-27B", { left: 94, top: 191, width: 360, height: 26 }, {
    fontSize: 15,
    bold: true,
    color: C.green,
  });
  text(slide, "Initial plan, then at most one repair or truncation recovery", { left: 94, top: 217, width: 740, height: 26 }, {
    fontSize: 21,
    color: C.white,
  });
  text(slide, "CONCURRENCY 4", { left: 955, top: 199, width: 226, height: 34 }, {
    fontSize: 16,
    bold: true,
    color: C.white,
    alignment: "right",
    verticalAlignment: "middle",
  });

  const routes = [
    {
      left: 64,
      label: "TARGET CELLS",
      heading: "Typed workbook operations",
      body: "Cell-level tasks can use only the declarative operation route. Every edit carries an explicit sheet and address.",
      band: C.green,
      footer: "Containment and formula checks before write",
    },
    {
      left: 656,
      label: "LARGER SHEET TASKS",
      heading: "Typed operations or screened Python",
      body: "Sheet-level transformations can use the same operations or screened Python with limited imports and resources.",
      band: C.blue,
      footer: "Temporary execution before publication",
    },
  ];
  for (const route of routes) {
    rect(slide, { left: route.left, top: 292, width: 560, height: 263 }, {
      fill: C.paper,
      line: { style: "solid", fill: C.line, width: 1 },
    });
    rect(slide, { left: route.left, top: 292, width: 8, height: 263 }, {
      fill: route.band,
      line: { fill: "none", width: 0 },
    });
    text(slide, route.label, { left: route.left + 34, top: 319, width: 480, height: 24 }, {
      fontSize: 13,
      bold: true,
      color: route.band,
    });
    text(slide, route.heading, { left: route.left + 34, top: 357, width: 485, height: 70 }, {
      fontSize: 29,
      bold: true,
      color: C.ink,
      lineSpacing: 1.03,
    });
    text(slide, route.body, { left: route.left + 34, top: 438, width: 485, height: 74 }, {
      fontSize: 18,
      color: C.muted,
      lineSpacing: 1.14,
    });
    text(slide, route.footer, { left: route.left + 34, top: 523, width: 485, height: 24 }, {
      fontSize: 14,
      bold: true,
      color: C.ink,
    });
  }
  rect(slide, { left: 64, top: 592, width: 1152, height: 55 }, {
    fill: C.paleGold,
    line: { style: "solid", fill: C.gold, width: 1 },
  });
  text(slide, "The second-call allowance is exclusive. A task never gets both repair and truncation recovery.", { left: 90, top: 607, width: 1098, height: 26 }, {
    fontSize: 19,
    bold: true,
    color: C.ink,
    verticalAlignment: "middle",
  });
  footer(slide, 3);
  notes(slide, [
    "Sources:",
    "exactsource/config.py",
    "exactsource/runner.py",
    "exactsource/contracts.py",
    "SUBMISSION.md section Current v2 model path",
    "FormulaBench v2 fixes the model at Qwen/Qwen3.8-27B and the concurrency at four.",
    "Each task has at most two logical model calls: initial plus one repair, or initial plus one cell truncation recovery.",
  ]);
}

// 4. Failure modes
{
  const slide = deck.slides.add();
  slide.background.fill = C.offWhite;
  title(slide, "A valid workbook can still contain a wrong answer");
  text(slide, "V2 contains failures it can verify. The organiser evaluator still decides correctness.", { left: 64, top: 108, width: 1110, height: 38 }, {
    fontSize: 23,
    color: C.muted,
  });
  const items = [
    ["01", "Context", "The needed sheet, range or existing formula is missing from the bounded context.", C.green],
    ["02", "Coverage", "The plan omits a target, duplicates an edit or writes outside the allowed area.", C.green],
    ["03", "Workbook", "The plan violates formula policy, resource limits or the save and reopen contract.", C.green],
    ["04", "Semantics", "The output opens and recalculates, yet the investment result is still wrong.", C.red],
  ];
  const lefts = [64, 356, 648, 940];
  for (let i = 0; i < items.length; i += 1) {
    const [number, heading, body, colour] = items[i];
    text(slide, number, { left: lefts[i], top: 205, width: 90, height: 52 }, {
      fontSize: 38,
      bold: true,
      color: colour,
    });
    text(slide, heading, { left: lefts[i], top: 278, width: 250, height: 42 }, {
      fontSize: 27,
      bold: true,
    });
    text(slide, body, { left: lefts[i], top: 337, width: 245, height: 150 }, {
      fontSize: 19,
      color: C.muted,
      lineSpacing: 1.14,
    });
    if (i < items.length - 1) {
      slide.shapes.add({
        geometry: "line",
        position: { left: lefts[i] + 270, top: 202, width: 0, height: 296 },
        fill: "none",
        line: { style: "solid", fill: C.line, width: 1 },
      });
    }
  }
  rect(slide, { left: 64, top: 545, width: 1152, height: 82 }, {
    fill: C.navy,
    line: { fill: "none", width: 0 },
  });
  text(slide, "Execution checks can reject unsafe work. They cannot certify a financially correct answer.", { left: 94, top: 566, width: 1092, height: 40 }, {
    fontSize: 23,
    bold: true,
    color: C.white,
    verticalAlignment: "middle",
  });
  footer(slide, 4);
  notes(slide, [
    "Sources:",
    "ARCHITECTURE.md",
    "exactsource/context.py",
    "exactsource/plans.py",
    "exactsource/sandbox.py",
    "The distinction between executable output and benchmark correctness is central to the organiser evaluator.",
  ]);
}

// 5. Workflow
{
  const slide = deck.slides.add();
  slide.background.fill = C.navy;
  title(slide, "V2 publishes only after the complete guarded path", true);
  text(slide, "The pristine workbook remains the recovery source throughout the task.", { left: 64, top: 108, width: 1050, height: 38 }, {
    fontSize: 23,
    color: "#AAB6C5",
  });

  const stages = [
    ["01", "Load", "Pristine input"],
    ["02", "Context", "Bounded evidence"],
    ["03", "Plan", "Initial call"],
    ["04", "Gate", "Parse and check"],
    ["05", "Recover", "Optional second call"],
    ["06", "Execute", "Temporary copy"],
    ["07", "Publish", "Validated output"],
  ];
  const gap = 14;
  const width = (1152 - gap * 6) / 7;
  for (let i = 0; i < stages.length; i += 1) {
    const left = 64 + i * (width + gap);
    const [number, heading, detail] = stages[i];
    rect(slide, { left, top: 194, width, height: 232 }, {
      fill: i === 4 ? "#1A304B" : C.navy2,
      line: { style: "solid", fill: i === 4 ? C.gold : C.lineDark, width: i === 4 ? 2 : 1 },
    });
    text(slide, number, { left: left + 18, top: 215, width: width - 36, height: 32 }, {
      fontSize: 15,
      bold: true,
      color: i === 4 ? C.gold : C.green,
    });
    text(slide, heading, { left: left + 18, top: 272, width: width - 36, height: 58 }, {
      fontSize: 23,
      bold: true,
      color: C.white,
      alignment: "center",
      verticalAlignment: "middle",
    });
    text(slide, detail, { left: left + 16, top: 350, width: width - 32, height: 50 }, {
      fontSize: 15,
      color: "#AAB6C5",
      alignment: "center",
      lineSpacing: 1.08,
    });
  }

  rect(slide, { left: 64, top: 470, width: 730, height: 140 }, {
    fill: "#13243B",
    line: { style: "solid", fill: C.lineDark, width: 1 },
  });
  text(slide, "PASS PATH", { left: 92, top: 493, width: 160, height: 24 }, {
    fontSize: 13,
    bold: true,
    color: C.green,
  });
  text(slide, "Validate targets, formulas, resources and workbook save. Then publish atomically.", { left: 92, top: 528, width: 670, height: 58 }, {
    fontSize: 21,
    color: C.white,
    lineSpacing: 1.12,
  });

  rect(slide, { left: 824, top: 470, width: 392, height: 140 }, {
    fill: "#2B2130",
    line: { style: "solid", fill: C.red, width: 1 },
  });
  text(slide, "REJECT PATH", { left: 852, top: 493, width: 160, height: 24 }, {
    fontSize: 13,
    bold: true,
    color: C.red,
  });
  text(slide, "Discard the work and copy the pristine input to the submission output.", { left: 852, top: 528, width: 336, height: 58 }, {
    fontSize: 19,
    color: C.white,
    lineSpacing: 1.12,
  });
  footer(slide, 5, true);
  notes(slide, [
    "Sources:",
    "exactsource/runner.py",
    "exactsource/artifacts.py",
    "exactsource/plans.py",
    "formulabench/v2.py",
    "Rejected work is never applied to the published submission workbook. The pristine input is retained as the fallback.",
  ]);
}

// 6. Publication contract
{
  const slide = deck.slides.add();
  slide.background.fill = C.offWhite;
  title(slide, "Checks before a workbook can be published");
  text(slide, "Typed operations and the sheet route share a fail-closed publication boundary.", { left: 64, top: 108, width: 1080, height: 38 }, {
    fontSize: 23,
    color: C.muted,
  });
  const rows = [
    ["Boundary", "Required evidence", "On failure"],
    ["Target containment", "Every edit remains inside the declared worksheet and target area.", "Reject the plan"],
    ["Resource policy", "The selected route stays within bounded time, memory and import limits.", "Stop execution"],
    ["Formula policy", "Formula assignments pass the workbook formula checks before publication.", "Reject the plan"],
    ["Workbook save", "A fresh copy saves successfully and reopens under the output contract.", "Discard the copy"],
    ["Output integrity", "The published workbook and recorded prediction agree on the task outcome.", "Use pristine fallback"],
  ];
  const table = slide.tables.add({
    rows: rows.length,
    columns: rows[0].length,
    left: 64,
    top: 170,
    width: 1152,
    height: 408,
    columnWidths: [225, 650, 277],
    values: tableRuns(rows),
  });
  styleTable(table, rows.length, rows[0].length, { boldFirstColumn: true });
  table.rows[0].height = 48;
  for (let row = 1; row < rows.length; row += 1) {
    table.rows[row].height = 72;
    table.getCell(row, 2).fill = C.mint;
    table.getCell(row, 2).text.style = {
      typeface: FONT,
      fontSize: 17,
      bold: true,
      color: C.greenDark,
      verticalAlignment: "middle",
      autoFit: "shrinkText",
      insets: { top: 7, right: 9, bottom: 7, left: 9 },
    };
  }
  text(slide, "Rejected plans preserve the pristine workbook.", { left: 64, top: 610, width: 1080, height: 36 }, {
    fontSize: 23,
    bold: true,
    color: C.ink,
  });
  footer(slide, 6);
  notes(slide, [
    "Sources:",
    "exactsource/plans.py",
    "exactsource/formula_safety.py",
    "exactsource/sandbox.py",
    "exactsource/artifacts.py",
    "The table describes deterministic guards. It is not evidence of formula correctness.",
  ]);
}

// 7. Migration parity
{
  const slide = deck.slides.add();
  slide.background.fill = C.paper;
  title(slide, "Migration parity reproduced all 400 retained outcomes");
  text(slide, "A credential-free replay compared the FormulaBench v2 wrapper with retained ExactSource artefacts.", { left: 64, top: 108, width: 1120, height: 38 }, {
    fontSize: 23,
    color: C.muted,
  });
  metric(slide, "400", "tasks checked", 64, 184, C.ink, 220);
  metric(slide, "369", "accepted plan replays", 330, 184, C.green, 250);
  metric(slide, "31", "pristine fallback checks", 630, 184, C.gold, 250);
  metric(slide, "0", "wrapper mismatches", 950, 184, C.blue, 250);

  rule(slide, 64, 314, 1152, C.line, 1);
  text(slide, "88.13 seconds", { left: 64, top: 350, width: 420, height: 64 }, {
    fontSize: 47,
    bold: true,
    color: C.ink,
  });
  text(slide, "host wall time, including container startup", { left: 68, top: 414, width: 440, height: 28 }, {
    fontSize: 17,
    color: C.muted,
  });
  text(slide, "Accepted plans", { left: 570, top: 354, width: 230, height: 28 }, {
    fontSize: 16,
    bold: true,
    color: C.green,
  });
  text(slide, "Rebuilt workbook content and compared the OOXML member inventory and payloads after normalising core timestamps.", { left: 570, top: 392, width: 590, height: 72 }, {
    fontSize: 19,
    color: C.ink,
    lineSpacing: 1.13,
  });
  text(slide, "Fallback plans", { left: 570, top: 488, width: 230, height: 28 }, {
    fontSize: 16,
    bold: true,
    color: C.gold,
  });
  text(slide, "Required byte-identical output against each pristine input workbook.", { left: 570, top: 526, width: 590, height: 50 }, {
    fontSize: 19,
    color: C.ink,
  });
  rect(slide, { left: 64, top: 600, width: 1152, height: 48 }, {
    fill: C.navy,
    line: { fill: "none", width: 0 },
  });
  text(slide, "Zero model calls. No golden workbooks read. This is migration parity, not benchmark accuracy.", { left: 90, top: 612, width: 1098, height: 25 }, {
    fontSize: 19,
    bold: true,
    color: C.white,
    verticalAlignment: "middle",
  });
  footer(slide, 7);
  notes(slide, [
    "Sources:",
    "experiments/v2_migration_validation.json section credential_free_parity",
    "tools/verify_v2_parity.py",
    "Recorded result: 400 checked, 369 accepted plan replays, 31 safe fallback byte checks, zero mismatches, 88.13 seconds including container startup.",
    "The replay made zero model calls, read no golden workbooks and does not establish FormulaBench v2 benchmark accuracy.",
  ]);
}

// 8. Evaluator path
{
  const slide = deck.slides.add();
  slide.background.fill = C.offWhite;
  title(slide, "The evaluator grades results, not plausible output");
  text(slide, "A complete v2 score requires the unchanged organiser evaluation path.", { left: 64, top: 108, width: 1030, height: 38 }, {
    fontSize: 23,
    color: C.muted,
  });
  const rows = [
    ["Stage", "Input", "What it establishes"],
    ["1  Recalculate", "Generated workbook opened by LibreOffice", "Formulas produce current workbook values"],
    ["2  Compare", "Every declared target cell", "Exact cell matches and mismatches"],
    ["3  Grade task", "Complete target set", "A pass only when the whole task is correct"],
    ["4  Aggregate", "All 400 evaluated tasks", "Task pass rate and cell accuracy"],
  ];
  const table = slide.tables.add({
    rows: rows.length,
    columns: rows[0].length,
    left: 64,
    top: 176,
    width: 1152,
    height: 350,
    columnWidths: [230, 400, 522],
    values: tableRuns(rows),
  });
  styleTable(table, rows.length, rows[0].length, { boldFirstColumn: true });
  table.rows[0].height = 50;
  for (let row = 1; row < rows.length; row += 1) {
    table.rows[row].height = 75;
    table.getCell(row, 0).fill = C.paleBlue;
    table.getCell(row, 0).text.style = {
      typeface: FONT,
      fontSize: 17,
      bold: true,
      color: C.ink,
      verticalAlignment: "middle",
      autoFit: "shrinkText",
      insets: { top: 7, right: 9, bottom: 7, left: 9 },
    };
  }
  rect(slide, { left: 64, top: 570, width: 1152, height: 76 }, {
    fill: C.navy,
    line: { fill: "none", width: 0 },
  });
  text(slide, "Opening, saving and recalculating are necessary checks. They are not a substitute for target-cell correctness.", { left: 92, top: 589, width: 1095, height: 40 }, {
    fontSize: 21,
    bold: true,
    color: C.white,
    verticalAlignment: "middle",
  });
  footer(slide, 8);
  notes(slide, [
    "Sources:",
    "evaluate.py",
    "SUBMISSION.md section Historical v1 verified evaluation",
    "experiments/v2_migration_validation.json",
    "The organiser-supplied evaluator is the benchmark authority for the public task set.",
  ]);
}

// 9. Route canaries
{
  const slide = deck.slides.add();
  slide.background.fill = C.offWhite;
  title(slide, "Two selected canaries exercised both execution routes");
  text(slide, "Each case ran through LibreOffice and the unchanged organiser evaluator.", { left: 64, top: 108, width: 1050, height: 38 }, {
    fontSize: 23,
    color: C.muted,
  });
  const canaries = [
    {
      left: 64,
      colour: C.green,
      label: "TYPED OPERATIONS",
      task: "Task 54513",
      value: "1 / 1",
      unit: "target cell correct",
      body: "Cell-level manipulation\nOne model call\nRuntime status: ok",
    },
    {
      left: 656,
      colour: C.blue,
      label: "SCREENED PYTHON",
      task: "Task 23-24",
      value: "5,510 / 5,510",
      unit: "target cells correct",
      body: "Sheet-level manipulation\nOne model call\nRuntime status: ok",
    },
  ];
  for (const item of canaries) {
    rect(slide, { left: item.left, top: 178, width: 560, height: 355 }, {
      fill: C.paper,
      line: { style: "solid", fill: C.line, width: 1 },
    });
    rect(slide, { left: item.left, top: 178, width: 560, height: 8 }, {
      fill: item.colour,
      line: { fill: "none", width: 0 },
    });
    text(slide, item.label, { left: item.left + 30, top: 210, width: 480, height: 24 }, {
      fontSize: 13,
      bold: true,
      color: item.colour,
    });
    text(slide, item.task, { left: item.left + 30, top: 250, width: 480, height: 42 }, {
      fontSize: 27,
      bold: true,
      color: C.ink,
    });
    text(slide, item.value, { left: item.left + 30, top: 314, width: 500, height: 72 }, {
      fontSize: item.value.length > 7 ? 47 : 57,
      bold: true,
      color: C.ink,
    });
    text(slide, item.unit, { left: item.left + 34, top: 387, width: 470, height: 30 }, {
      fontSize: 18,
      bold: true,
      color: item.colour,
    });
    rule(slide, item.left + 30, 438, 500, C.line, 1);
    text(slide, item.body, { left: item.left + 30, top: 456, width: 500, height: 62 }, {
      fontSize: 16,
      color: C.muted,
      lineSpacing: 1.17,
    });
  }
  rect(slide, { left: 64, top: 574, width: 1152, height: 74 }, {
    fill: C.paleRed,
    line: { style: "solid", fill: C.red, width: 1 },
  });
  text(slide, "These are selected route diagnostics. A two-case pass count cannot estimate performance on the 400-task benchmark.", { left: 90, top: 592, width: 1095, height: 40 }, {
    fontSize: 20,
    bold: true,
    color: C.ink,
    verticalAlignment: "middle",
  });
  footer(slide, 9);
  notes(slide, [
    "Sources:",
    "experiments/v2_migration_validation.json section live_canaries",
    "Task 54513 used typed operations and passed one of one target cell.",
    "Task 23-24 used screened Python and passed 5,510 of 5,510 target cells.",
    "Both cases were selected as route diagnostics and do not estimate the 400-task pass rate.",
  ]);
}

// 10. Score record
{
  const slide = deck.slides.add();
  slide.background.fill = C.paper;
  title(slide, "Historical scores stay attached to their recorded runs");
  text(slide, "No earlier result is transferred to FormulaBench v2.", { left: 64, top: 108, width: 1030, height: 38 }, {
    fontSize: 23,
    color: C.muted,
  });

  const chart = slide.charts.add("bar", {
    position: { left: 64, top: 180, width: 720, height: 335 },
    categories: ["FormulaBench v1", "ExactSource source run"],
    series: [{
      name: "Task pass rate",
      values: [0.3325, 0.755],
      valuesFormatCode: "0.00%",
      fill: C.green,
      points: [
        { idx: 0, fill: C.blue },
        { idx: 1, fill: C.green },
      ],
    }],
    barOptions: { direction: "bar", grouping: "clustered", gapWidth: 58 },
    hasLegend: false,
    xAxis: {
      visible: false,
      textStyle: { typeface: FONT, fontSize: 15, fill: C.muted },
      majorGridlines: null,
      line: { fill: "none", width: 0 },
    },
    yAxis: {
      visible: true,
      min: 0,
      max: 1,
      majorUnit: 0.2,
      numberFormatCode: "0%",
      textStyle: { typeface: FONT, fontSize: 17, fill: C.ink, bold: true },
      majorGridlines: null,
      line: { fill: "none", width: 0 },
    },
    dataLabels: {
      showValue: true,
      position: "outEnd",
      textStyle: { typeface: FONT, fontSize: 18, fill: C.ink, bold: true },
    },
    chartFill: C.paper,
    chartLine: { fill: "none", width: 0 },
    plotAreaFill: C.paper,
    plotAreaLine: { fill: "none", width: 0 },
  });
  chart.fontFamily = FONT;

  rect(slide, { left: 842, top: 180, width: 374, height: 335 }, {
    fill: C.navy,
    line: { fill: "none", width: 0 },
  });
  text(slide, "FORMULABENCH V2", { left: 875, top: 214, width: 305, height: 26 }, {
    fontSize: 14,
    bold: true,
    color: C.green,
  });
  text(slide, "UNSCORED", { left: 873, top: 276, width: 310, height: 66 }, {
    fontSize: 50,
    bold: true,
    color: C.white,
  });
  text(slide, "No complete 400-task result had been published at deck capture.", { left: 875, top: 365, width: 300, height: 86 }, {
    fontSize: 20,
    color: "#C6D2DF",
    lineSpacing: 1.14,
  });

  const rows = [
    ["Recorded run", "Task result", "Cell result"],
    ["FormulaBench v1", "133 / 400   33.25%", "40.56%"],
    ["ExactSource source run", "302 / 400   75.50%", "80.06%"],
  ];
  const table = slide.tables.add({
    rows: rows.length,
    columns: rows[0].length,
    left: 64,
    top: 548,
    width: 1152,
    height: 100,
    columnWidths: [425, 420, 307],
    values: tableRuns(rows),
  });
  styleTable(table, rows.length, rows[0].length, { boldFirstColumn: true });
  table.rows[0].height = 32;
  table.rows[1].height = 34;
  table.rows[2].height = 34;
  footer(slide, 10);
  notes(slide, [
    "Sources:",
    "submissions/formulabench/results.json",
    "experiments/v2_migration_validation.json sections upstream_scored_reference and status",
    "FormulaBench v1: 133 of 400 tasks, 33.25 per cent pass rate and 40.56 per cent cell accuracy.",
    "Historical ExactSource source run: 302 of 400 tasks, 75.50 per cent pass rate and 80.06 per cent cell accuracy.",
    "The ExactSource result is the source baseline whose implementation was migrated. It is not a FormulaBench v2 result.",
  ]);
}

// 11. Reproduce and inspect
{
  const slide = deck.slides.add();
  slide.background.fill = C.navy;
  title(slide, "Run it, inspect it, then score the complete output", true);
  text(slide, "The public repository keeps the runtime, parity record and historical score distinct.", { left: 64, top: 108, width: 1100, height: 38 }, {
    fontSize: 23,
    color: "#AAB6C5",
  });

  text(slide, "RUN V2", { left: 64, top: 184, width: 200, height: 25 }, {
    fontSize: 14,
    bold: true,
    color: C.green,
  });
  rect(slide, { left: 64, top: 220, width: 730, height: 130 }, {
    fill: "#13243B",
    line: { style: "solid", fill: C.lineDark, width: 1 },
  });
  text(slide, "uv run python data/download.py\n./scripts/run_docker.sh data/spreadsheetbench_verified_400 submissions/formulabench-v2", { left: 88, top: 246, width: 680, height: 78 }, {
    typeface: MONO,
    fontSize: 17,
    color: C.white,
    lineSpacing: 1.24,
  });
  text(slide, "Set TINKER_API_KEY privately in the shell before a paid run.", { left: 64, top: 365, width: 720, height: 26 }, {
    fontSize: 16,
    color: "#AAB6C5",
  });

  text(slide, "EVIDENCE ANCHORS", { left: 64, top: 431, width: 250, height: 25 }, {
    fontSize: 14,
    bold: true,
    color: C.green,
  });
  text(slide, "experiments/v2_migration_validation.json\nsubmissions/formulabench/results.json\nSUBMISSION.md", { left: 64, top: 472, width: 570, height: 96 }, {
    typeface: MONO,
    fontSize: 18,
    color: "#C6D2DF",
    lineSpacing: 1.3,
  });

  rect(slide, { left: 842, top: 184, width: 374, height: 384 }, {
    fill: C.navy2,
    line: { style: "solid", fill: C.lineDark, width: 1 },
  });
  text(slide, "CAPTURE STATUS", { left: 872, top: 214, width: 300, height: 25 }, {
    fontSize: 14,
    bold: true,
    color: C.gold,
  });
  text(slide, "Current default", { left: 872, top: 268, width: 300, height: 38 }, {
    fontSize: 27,
    bold: true,
    color: C.white,
  });
  text(slide, "FormulaBench v2", { left: 872, top: 310, width: 300, height: 32 }, {
    fontSize: 20,
    color: "#C6D2DF",
  });
  rule(slide, 872, 371, 314, C.lineDark, 1);
  text(slide, "Benchmark state", { left: 872, top: 398, width: 300, height: 28 }, {
    fontSize: 15,
    bold: true,
    color: C.green,
  });
  text(slide, "Unscored at capture", { left: 872, top: 438, width: 300, height: 38 }, {
    fontSize: 26,
    bold: true,
    color: C.white,
  });
  text(slide, "6 September 2026", { left: 872, top: 493, width: 300, height: 26 }, {
    fontSize: 16,
    color: "#AAB6C5",
  });

  text(slide, "PUBLIC REPOSITORY", { left: 64, top: 606, width: 240, height: 22 }, {
    fontSize: 13,
    bold: true,
    color: C.green,
  });
  text(slide, "github.com/MasteraSnackin/FormulaBench", { left: 310, top: 604, width: 480, height: 26 }, {
    fontSize: 20,
    bold: true,
    color: C.white,
    link: "https://github.com/MasteraSnackin/FormulaBench",
  });
  text(slide, "Private benchmark remains unseen", { left: 842, top: 606, width: 374, height: 22 }, {
    fontSize: 15,
    color: "#AAB6C5",
    alignment: "right",
  });
  footer(slide, 11, true);
  notes(slide, [
    "Sources:",
    "README.md sections Quick start and V2 migration evidence",
    "SUBMISSION.md sections Version and evidence status and Code",
    "scripts/run_docker.sh",
    "https://github.com/MasteraSnackin/FormulaBench",
    "Capture status is dated because a later complete organiser-evaluated run may supersede it.",
  ]);
}

await fs.mkdir(path.dirname(outputPath), { recursive: true });
const inspection = await deck.inspect({
  kind: "slide,textbox,shape,table,chart,notes,layout",
  maxChars: 100000,
});
await fs.writeFile(inspectPath, inspection.ndjson);
await (await PresentationFile.exportPptx(deck)).save(outputPath);

console.log(JSON.stringify({ outputPath, inspectPath, slideCount: 11 }, null, 2));
