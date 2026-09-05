"""Prompt text and exact authored-context budgeting."""

from __future__ import annotations

import json
from collections.abc import Iterable
from typing import Protocol

from openpyxl.utils.cell import range_boundaries

from .constants import MAX_AUTHORED_CHARS

SYSTEM_PROMPT = (
    "You are a spreadsheet expert. Use only the supplied instruction and workbook evidence. "
    "Return the final content or an explicit preserve instruction for every answer cell. Use exact "
    "worksheet names and A1 addresses. A value may be a string, number, boolean, null, a typed "
    "Excel date, or an Excel formula beginning with '='. "
    "Use a compact fill for a large rectangular result: its value is anchored at the top-left, "
    "formula references are copied relatively as in Excel, and constants repeat unchanged. "
    "Prefer formulas when requested or when results must remain live. Use only listed target "
    "sheets, including a target explicitly marked for creation; never introduce any other sheet, "
    "target, source value, merged range, or missing range. Call submit_spreadsheet_answer "
    "exactly once with the completed answer object and do not call any other tool. If tool calls "
    "are unavailable, return that same object as strict JSON text only, without Markdown or "
    "explanation."
)

RESPONSE_FORMAT = (
    '{"cells":[],"fills":[{"sheet":"Sheet1","range":"B6:B100",'
    '"value":"=A6*2"}],"preserve_ranges":[{"sheet":"Sheet1",'
    '"range":"A1:A5"}],"unmerge_ranges":[]}'
)


class TargetLike(Protocol):
    sheet: str
    cell_range: str
    dynamic: bool


def target_summary_lines(ranges: Iterable[TargetLike]) -> list[str]:
    """Render target objects without coupling prompts to a concrete dataclass."""

    lines: list[str] = []
    for target in ranges:
        notes: list[str] = []
        if target.dynamic:
            notes.append("row extent must be inferred")
        else:
            min_col, min_row, max_col, max_row = range_boundaries(target.cell_range)
            if None in {
                min_col,
                min_row,
                max_col,
                max_row,
            }:  # pragma: no cover - contract invariant
                raise ValueError("finite target has no finite cell count")
            cell_count = (max_col - min_col + 1) * (max_row - min_row + 1)
            notes.append(f"exactly {cell_count:,} cells; include unchanged cells and blanks")
        resolution = getattr(target, "resolution", None)
        if hasattr(resolution, "value"):
            resolution = resolution.value
        if resolution == "created":
            notes.append("create this exact target sheet")
        suffix = f" ({'; '.join(notes)})" if notes else ""
        lines.append(f"- {json.dumps(target.sheet, ensure_ascii=True)}!{target.cell_range}{suffix}")
    return lines


def prompt_scaffold(*, instruction: str, target_lines: Iterable[str]) -> tuple[str, str]:
    prefix = f"## Instruction\n{instruction}\n\n## Workbook evidence\n"
    suffix = (
        "\n\n## Required answer targets\n"
        + "\n".join(target_lines)
        + "\n\nFinite targets are complete partitions, not patches. Cover every target cell "
        "exactly "
        "once with cells, fill-expanded values, or preserve_ranges. Repeat unchanged template or "
        "example values unless preserve_ranges explicitly keeps them; represent every blank that "
        "must remain or become blank with null or a null fill. One omitted, overlapping, "
        "duplicate, or outside cell rejects the whole answer. For a target with an inferred row "
        "extent, return "
        "a dense finite rectangle through at least that worksheet's shown final row, extending it "
        "if any answer on that sheet reaches a lower row. A fill or preserve range must have two "
        "finite A1 endpoints (for example B6:B100), stay within one listed target, and not overlap "
        "another returned cell or range. A fill formula is written for the top-left cell and "
        "copied "
        "relatively; a constant repeats. Use preserve_ranges only for cells whose current workbook "
        'contents must remain unchanged. The cells array is always required: use "cells":[] when '
        "ranges contain all answers. For a real Excel date, use a typed value such as "
        '{"type":"date","value":"2024-02-22"} or '
        '{"type":"datetime","value":"2024-02-22T00:00:00"}, or a formula such as '
        '"=DATE(2024,2,22)"; never use date-looking text. Formula strings must not exceed 8,192 '
        "characters; datetime fractions must resolve exactly to milliseconds; other cell strings "
        "must not exceed 32,767. When grouping repeated sections, "
        "keep each source section separate unless instructed otherwise, emit each requested key "
        "once, preserve the requested sort order, and reconcile group values and totals. Prefer "
        "formulas over mental arithmetic when practical. When retaining a shown merged range, "
        "return its value at the anchor (top-left) cell and null for every targeted non-anchor "
        "cell, "
        "or explicitly preserve the complete merged range. Use unmerge_ranges only when the "
        "instruction requires unmerging; each entry must name an exact existing merged range shown "
        "in the workbook evidence, on an exact sheet, and intersect an answer target. Otherwise "
        "omit it or use an empty array. Before submitting, silently verify that cells plus "
        "expanded "
        "fills and preserves match every target's stated count. Do not return cells outside the "
        "listed sheets and ranges.\n\n"
        "## Response schema\n" + RESPONSE_FORMAT
    )
    return prefix, suffix


def workbook_char_budget(*, prefix: str, suffix: str) -> int:
    fixed = len(SYSTEM_PROMPT) + len(prefix) + len(suffix)
    budget = MAX_AUTHORED_CHARS - fixed
    if budget < 256:
        raise ValueError(
            f"instruction and response contract leave only {budget} workbook characters"
        )
    return budget


def assert_authored_budget(user_prompt: str) -> None:
    authored = len(SYSTEM_PROMPT) + len(user_prompt)
    if authored > MAX_AUTHORED_CHARS:
        raise ValueError(
            f"authored prompt is {authored} characters; maximum is {MAX_AUTHORED_CHARS}"
        )
