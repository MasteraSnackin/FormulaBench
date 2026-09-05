from __future__ import annotations

from types import SimpleNamespace

import pytest

from formulabench.constants import MAX_AUTHORED_CHARS
from formulabench.prompts import (
    RESPONSE_FORMAT,
    SYSTEM_PROMPT,
    assert_authored_budget,
    prompt_scaffold,
    target_summary_lines,
    workbook_char_budget,
)


def test_workbook_budget_accounts_for_every_authored_character() -> None:
    prefix, suffix = prompt_scaffold(instruction="Fill B2.", target_lines=['- "Sheet1"!B2'])
    budget = workbook_char_budget(prefix=prefix, suffix=suffix)
    workbook = "x" * budget
    user_prompt = prefix + workbook + suffix
    assert len(SYSTEM_PROMPT) + len(user_prompt) == MAX_AUTHORED_CHARS
    assert_authored_budget(user_prompt)


def test_budget_rejects_an_instruction_that_crowds_out_workbook_evidence() -> None:
    prefix, suffix = prompt_scaffold(
        instruction="x" * MAX_AUTHORED_CHARS,
        target_lines=['- "Sheet1"!B2'],
    )
    with pytest.raises(ValueError, match="workbook characters"):
        workbook_char_budget(prefix=prefix, suffix=suffix)


def test_prompt_defines_fills_only_and_exact_unmerge_contract() -> None:
    _, suffix = prompt_scaffold(
        instruction="Fill the result and unmerge the instructed range.",
        target_lines=['- "Sheet1"!B2:B100'],
    )

    assert '"cells":[]' in RESPONSE_FORMAT
    assert '"fills"' in RESPONSE_FORMAT
    assert '"preserve_ranges"' in RESPONSE_FORMAT
    assert '"unmerge_ranges"' in RESPONSE_FORMAT
    assert "top-left cell and copied relatively" in suffix
    assert 'use "cells":[] when ranges contain all answers' in suffix
    assert "complete partitions, not patches" in suffix
    assert "One omitted, overlapping, duplicate, or outside cell rejects the whole answer" in suffix
    assert "never use date-looking text" in suffix
    assert "keep each source section separate" in suffix
    assert "silently verify that cells plus expanded fills and preserves" in suffix
    assert "through at least that worksheet's shown final row" in suffix
    assert "Formula strings must not exceed 8,192 characters" in suffix
    assert "return its value at the anchor" in suffix
    assert "null for every targeted non-anchor cell" in suffix
    assert "exact existing merged range shown in the workbook evidence" in suffix


def test_system_prompt_requests_exactly_one_answer_tool_with_json_fallback() -> None:
    assert "Call submit_spreadsheet_answer exactly once" in SYSTEM_PROMPT
    assert "do not call any other tool" in SYSTEM_PROMPT
    assert "If tool calls are unavailable" in SYSTEM_PROMPT
    assert "strict JSON text only" in SYSTEM_PROMPT


def test_target_summary_marks_only_explicitly_created_sheet() -> None:
    lines = target_summary_lines(
        [
            SimpleNamespace(
                sheet="New Output",
                cell_range="A1:A2",
                dynamic=False,
                resolution="created",
            ),
            SimpleNamespace(
                sheet="Existing",
                cell_range="B:B",
                dynamic=True,
                resolution="existing",
            ),
        ]
    )

    assert lines == [
        '- "New Output"!A1:A2 (exactly 2 cells; include unchanged cells and blanks; '
        "create this exact target sheet)",
        '- "Existing"!B:B (row extent must be inferred)',
    ]
