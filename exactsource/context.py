"""Deterministic, formula-aware prompt context for a SpreadsheetBench task."""

from __future__ import annotations

import hashlib
import json
import re
from dataclasses import dataclass
from datetime import date, datetime, time
from decimal import Decimal
from typing import Any

from exactsource.config import CONTEXT_CHAR_BUDGET
from exactsource.contracts import ContextPack, QualifiedRange, TaskSpec
from exactsource.ranges import (
    RangeSyntaxError,
    format_qualified_range,
    parse_qualified_ranges,
    range_bounds,
    range_cell_count,
)
from exactsource.workbook import CellSnapshot, WorkbookInspector

# Excel worksheet names cannot exceed 31 characters, so this parser default
# cannot collide with a real worksheet.  It lets us retain whether each parsed
# data_position range was explicitly qualified without changing the public
# range parser contract.
_UNQUALIFIED_DATA_SHEET = "__exactsource_unqualified_data_position__"

# These four sampled sections carry the workbook evidence needed to solve a
# task. On oversized contexts each receives up to this explicit reserve before
# spare capacity is shared evenly. Short sections return unused capacity to the
# others. The workbook manifest is kept whole whenever it independently fits
# within the caller's budget.
_SAMPLE_SECTION_RESERVE_CHARS = 6_000
_MANIFEST_CONTEXT_SECTION = "Workbook structure"
_SECTION_HEADER_RE = re.compile(r"(?m)^## (?P<name>[^\n]+)")
_CHILD_HEADER_RE = re.compile(r"(?m)^### [^\n]+")
_CHILD_STATUS_PREFIXES = (
    "- Graded cells=",
    "- Region cells=",
    "- The requested worksheet does not exist",
    "- Worksheet is absent from the input workbook.",
    "- No additional target or neighbouring cell evidence.",
    "- All sampled populated cells already appear in higher-priority context.",
)

CellCoordinate = tuple[str, int, int]


@dataclass(frozen=True, slots=True)
class _ChildBlock:
    """A range- or worksheet-scoped evidence block inside a context section."""

    heading: str
    status: str
    body: str


def _json_value(value: Any) -> str:
    if isinstance(value, (datetime, date, time)):
        value = value.isoformat()
    elif isinstance(value, Decimal):
        value = str(value)
    return json.dumps(value, ensure_ascii=False, sort_keys=True, default=str)


def _cell_line(cell: CellSnapshot) -> str:
    if cell.formula is not None:
        content = f"formula={_json_value(cell.formula)}"
        if cell.cached_value is not None:
            content += f"; cached={_json_value(cell.cached_value)}"
        if cell.formula_ref is not None:
            content += f"; array_ref={_json_value(cell.formula_ref)}"
    else:
        content = f"value={_json_value(cell.value)}"
    details: list[str] = []
    if cell.number_format != "General":
        details.append(f"format={_json_value(cell.number_format)}")
    if cell.style_id:
        details.append(f"style={cell.style_id}")
    suffix = "" if not details else "; " + "; ".join(details)
    return f"- {cell.coordinate}: {content}{suffix}"


def _range_count(inspector: WorkbookInspector, target: QualifiedRange) -> int | None:
    try:
        worksheet = inspector.worksheet(target.sheet)
    except KeyError:
        return None
    return range_cell_count(
        target.cells,
        max_row=worksheet.max_row,
        max_column=worksheet.max_column,
    )


def _cell_coordinate(cell: CellSnapshot) -> CellCoordinate:
    """Return a workbook-qualified coordinate suitable for evidence de-duplication."""

    return cell.sheet, cell.row, cell.column


def _informative_blank(cell: CellSnapshot) -> bool:
    """Retain blank cells whose formatting still provides workbook evidence."""

    return cell.number_format != "General" or cell.style_id != 0


def _seen_on_sheet(seen: set[CellCoordinate], sheet: str) -> frozenset[tuple[int, int]]:
    return frozenset((row, column) for seen_sheet, row, column in seen if seen_sheet == sheet)


def _unique_cells(
    cells: tuple[CellSnapshot, ...],
    seen: set[CellCoordinate],
    *,
    retain_blank_coordinates: frozenset[CellCoordinate] = frozenset(),
) -> tuple[CellSnapshot, ...]:
    """Select new cell evidence while keeping the first, highest-priority occurrence.

    ``retain_blank_coordinates`` identifies sampled answer cells. Ordinary
    blank neighbours are omitted, but explicit answer cells remain visible
    even when they have neither a value nor formatting.
    """

    selected: list[CellSnapshot] = []
    for cell in cells:
        coordinate = _cell_coordinate(cell)
        keep_blank = coordinate in retain_blank_coordinates
        if cell.is_blank and not keep_blank and not _informative_blank(cell):
            continue
        if coordinate in seen:
            continue
        seen.add(coordinate)
        selected.append(cell)
    return tuple(selected)


def _append_target_context(
    lines: list[str],
    inspector: WorkbookInspector,
    targets: tuple[QualifiedRange, ...],
    seen: set[CellCoordinate],
) -> None:
    lines.extend(("", "## Answer-target context"))
    for target in targets:
        label = format_qualified_range(target)
        lines.extend(("", f"### {label}"))
        try:
            resolved = inspector.resolve_sheet(target.sheet)
        except KeyError:
            lines.append(
                "- The requested worksheet does not exist in the input workbook; "
                "the transformation may need to create it."
            )
            continue
        worksheet = inspector.worksheet(resolved)
        count = _range_count(inspector, target)
        min_col, min_row, max_col, max_row = range_bounds(
            target.cells,
            max_row=worksheet.max_row,
            max_column=worksheet.max_column,
        )

        resolution = (
            "" if resolved == target.sheet else f"; resolved worksheet={_json_value(resolved)}"
        )
        lines.append(f"- Graded cells={count}{resolution}")
        nearby = inspector.neighbourhood(target, row_margin=2, column_margin=2, limit=180)
        explicit_target_coordinates = frozenset(
            _cell_coordinate(cell)
            for cell in nearby
            if cell.sheet == resolved
            and min_row <= cell.row <= max_row
            and min_col <= cell.column <= max_col
        )
        unique_nearby = _unique_cells(
            nearby,
            seen,
            retain_blank_coordinates=explicit_target_coordinates,
        )
        if unique_nearby:
            lines.extend(_cell_line(cell) for cell in unique_nearby)
        else:
            lines.append("- No additional target or neighbouring cell evidence.")
        formulas = _unique_cells(
            inspector.formulas_near(
                target,
                row_margin=12,
                column_margin=6,
                limit=60,
                excluded=_seen_on_sheet(seen, resolved),
            ),
            seen,
        )
        if formulas:
            lines.append("- Nearby formula pattern cells:")
            lines.extend(f"  {_cell_line(cell)[2:]}" for cell in formulas)


def _instruction_mention_position(instruction: str, sheet: str) -> int | None:
    """Locate an explicit worksheet-name mention without substring matches."""

    haystack = instruction.casefold()
    needle = sheet.strip().casefold()
    if not needle:
        return None
    offset = 0
    while True:
        position = haystack.find(needle, offset)
        if position < 0:
            return None
        end = position + len(needle)
        left_ok = position == 0 or not (
            haystack[position - 1].isalnum() or haystack[position - 1] == "_"
        )
        right_ok = end == len(haystack) or not (haystack[end].isalnum() or haystack[end] == "_")
        if left_ok and right_ok:
            return position
        offset = position + 1


def _prioritised_sheet_names(inspector: WorkbookInspector, instruction: str) -> tuple[str, ...]:
    """Return every existing worksheet, with instruction mentions first."""

    ranked: list[tuple[bool, int, int, str]] = []
    worksheet_names = tuple(sheet.name for sheet in inspector.manifest().sheets)
    for workbook_index, sheet in enumerate(worksheet_names):
        position = _instruction_mention_position(instruction, sheet)
        ranked.append(
            (
                position is None,
                0 if position is None else position,
                workbook_index,
                sheet,
            )
        )
    ranked.sort()
    return tuple(item[3] for item in ranked)


def _data_ranges(
    task: TaskSpec,
    inspector: WorkbookInspector,
) -> tuple[tuple[QualifiedRange, ...], str | None]:
    if not task.data_position:
        return (), None
    default = task.answer_ranges[0].sheet if task.answer_ranges else None
    try:
        if not task.instruction_type.casefold().startswith("sheet"):
            return parse_qualified_ranges(task.data_position, default_sheet=default), None

        parsed = parse_qualified_ranges(
            task.data_position,
            default_sheet=_UNQUALIFIED_DATA_SHEET,
        )
        if all(source.sheet != _UNQUALIFIED_DATA_SHEET for source in parsed):
            return parsed, None

        candidates = _prioritised_sheet_names(inspector, task.instruction)
        expanded: list[QualifiedRange] = []
        for source in parsed:
            if source.sheet != _UNQUALIFIED_DATA_SHEET:
                expanded.append(source)
                continue
            expanded.extend(QualifiedRange(sheet=sheet, cells=source.cells) for sheet in candidates)
        return tuple(expanded), None
    except RangeSyntaxError as exc:
        return (), str(exc)


def _append_data_context(
    lines: list[str],
    inspector: WorkbookInspector,
    task: TaskSpec,
    seen: set[CellCoordinate],
) -> None:
    sources, error = _data_ranges(task, inspector)
    lines.extend(("", "## Declared source regions"))
    if error is not None:
        lines.append(f"- Could not parse data_position={_json_value(task.data_position)}: {error}")
        return
    if not sources:
        lines.append("- No data_position metadata was supplied.")
        return
    for source in sources:
        label = format_qualified_range(source)
        lines.extend(("", f"### {label}"))
        try:
            resolved = inspector.resolve_sheet(source.sheet)
        except KeyError:
            lines.append("- Worksheet is absent from the input workbook.")
            continue
        count = _range_count(inspector, source)
        resolution = (
            "" if resolved == source.sheet else f"; resolved worksheet={_json_value(resolved)}"
        )
        lines.append(f"- Region cells={count}{resolution}; deterministic sparse sample follows.")
        sample = _unique_cells(
            inspector.sample_range(
                resolved,
                source.cells,
                limit=180,
                include_blank=False,
                excluded=_seen_on_sheet(seen, resolved),
            ),
            seen,
        )
        if sample:
            lines.extend(_cell_line(cell) for cell in sample)
        else:
            lines.append("- All sampled populated cells already appear in higher-priority context.")


def _append_manifest(lines: list[str], inspector: WorkbookInspector) -> None:
    manifest = inspector.manifest()
    lines.extend(("", "## Workbook structure"))
    for sheet in manifest.sheets:
        lines.append(
            "- "
            f"sheet={_json_value(sheet.name)}; state={sheet.state}; "
            f"declared={sheet.declared_range}; effective={sheet.effective_range or 'empty'}; "
            f"materialised={sheet.materialised_cells}; nonempty={sheet.nonempty_cells}; "
            f"formulas={sheet.formula_cells}"
        )
        if sheet.freeze_panes:
            lines.append(f"  freeze_panes={sheet.freeze_panes}")
        if sheet.auto_filter:
            lines.append(f"  auto_filter={sheet.auto_filter}")
        if sheet.merged_ranges:
            shown = sheet.merged_ranges[:80]
            lines.append(f"  merged_ranges={_json_value(list(shown))}")
            if len(sheet.merged_ranges) > len(shown):
                lines.append(f"  merged_ranges_omitted={len(sheet.merged_ranges) - len(shown)}")
        for table in sheet.tables:
            lines.append(
                f"  table={_json_value(table.name)}; ref={table.ref}; "
                f"style={_json_value(table.style)}; "
                f"columns={_json_value(list(table.columns))}"
            )

    if manifest.defined_names:
        lines.append("- Defined names:")
        for name in manifest.defined_names[:120]:
            local = (
                "" if name.local_sheet is None else f"; local_sheet={_json_value(name.local_sheet)}"
            )
            lines.append(
                f"  name={_json_value(name.name)}; refers_to={_json_value(name.refers_to)}"
                f"{local}; hidden={str(name.hidden).lower()}"
            )
        if len(manifest.defined_names) > 120:
            lines.append(f"  defined_names_omitted={len(manifest.defined_names) - 120}")


def _append_formula_catalogue(
    lines: list[str],
    inspector: WorkbookInspector,
    seen: set[CellCoordinate],
) -> None:
    lines.extend(("", "## Workbook-wide formula patterns"))
    found_formula = False
    emitted_formula = False
    sheets_with_formulas = {
        sheet.name for sheet in inspector.manifest().sheets if sheet.formula_cells > 0
    }
    for sheet in inspector.sheet_names:
        if sheet in sheets_with_formulas:
            found_formula = True
        formulas = inspector.formula_sample(
            sheet,
            limit=60,
            excluded=_seen_on_sheet(seen, sheet),
        )
        if not formulas:
            continue
        formulas = _unique_cells(formulas, seen)
        if not formulas:
            continue
        emitted_formula = True
        lines.extend(("", f"### Worksheet {_json_value(sheet)}"))
        lines.extend(f"  {_cell_line(cell)[2:]}" for cell in formulas[:60])
    if not found_formula:
        lines.append("- No existing formulas were found in the sampled workbook cells.")
    elif not emitted_formula:
        lines.append("- All sampled formulas already appear in higher-priority context.")


def _append_sparse_samples(
    lines: list[str],
    inspector: WorkbookInspector,
    seen: set[CellCoordinate],
) -> None:
    lines.extend(("", "## Other populated workbook cells"))
    emitted = False
    for sheet in inspector.sheet_names:
        cells = _unique_cells(
            inspector.populated_sample(
                sheet,
                limit=90,
                excluded=_seen_on_sheet(seen, sheet),
            ),
            seen,
        )
        if not cells:
            continue
        emitted = True
        lines.extend(("", f"### Worksheet {_json_value(sheet)}"))
        lines.extend(_cell_line(cell) for cell in cells)
    if not emitted:
        lines.append("- No additional populated cells beyond higher-priority context.")


def _trim_section_lines(lines: list[str]) -> list[str]:
    start = 0
    end = len(lines)
    while start < end and not lines[start]:
        start += 1
    while end > start and not lines[end - 1]:
        end -= 1
    return lines[start:end]


def _render_context_sections(task: TaskSpec) -> tuple[tuple[str, str], ...]:
    target: list[str] = []
    manifest: list[str] = []
    sources: list[str] = []
    formulas: list[str] = []
    sparse: list[str] = []
    with WorkbookInspector(task.init_xlsx) as inspector:
        seen: set[CellCoordinate] = set()
        _append_target_context(target, inspector, task.answer_ranges, seen)
        _append_manifest(manifest, inspector)
        _append_data_context(sources, inspector, task, seen)
        _append_formula_catalogue(formulas, inspector, seen)
        _append_sparse_samples(sparse, inspector, seen)

    section_lines = (
        ("Answer-target context", target),
        ("Workbook structure", manifest),
        ("Declared source regions", sources),
        ("Workbook-wide formula patterns", formulas),
        ("Other populated workbook cells", sparse),
    )
    sections: list[tuple[str, str]] = []
    for index, (name, lines) in enumerate(section_lines):
        suffix = "\n" if index + 1 == len(section_lines) else "\n\n"
        sections.append((name, "\n".join(_trim_section_lines(lines)) + suffix))
    return tuple(sections)


def _render_full_context(task: TaskSpec) -> str:
    return "".join(text for _name, text in _render_context_sections(task))


def _context_sections(full_text: str) -> tuple[tuple[str, str], ...]:
    """Split rendered context at stable top-level Markdown section markers."""

    matches = tuple(_SECTION_HEADER_RE.finditer(full_text))
    if not matches:
        return (("__preamble__", full_text),)

    sections: list[tuple[str, str]] = []
    if matches[0].start() > 0:
        sections.append(("__preamble__", full_text[: matches[0].start()]))
    for index, match in enumerate(matches):
        end = matches[index + 1].start() if index + 1 < len(matches) else len(full_text)
        sections.append((match.group("name"), full_text[match.start() : end]))
    return tuple(sections)


def _fair_allocations(capacities: tuple[int, ...], budget: int) -> tuple[int, ...]:
    """Water-fill ``budget`` across capacities with stable index tie-breaking."""

    allocations = [0] * len(capacities)
    remaining = min(max(0, budget), sum(capacities))
    active = [index for index, capacity in enumerate(capacities) if capacity > 0]
    while active and remaining > 0:
        share = remaining // len(active)
        if share == 0:
            for index in active[:remaining]:
                allocations[index] += 1
            break

        saturated = [index for index in active if capacities[index] - allocations[index] <= share]
        if saturated:
            for index in saturated:
                grant = capacities[index] - allocations[index]
                allocations[index] += grant
                remaining -= grant
            active = [index for index in active if index not in saturated]
            continue

        for index in active:
            allocations[index] += share
        remaining -= share * len(active)
        for index in active[:remaining]:
            allocations[index] += 1
        break
    return tuple(allocations)


def _section_allocations(
    sections: tuple[tuple[str, str], ...],
    budget: int,
) -> tuple[int, ...]:
    """Protect the manifest and fairly divide all remaining section capacity."""

    lengths = tuple(len(text) for _name, text in sections)
    manifests = tuple(
        index for index, (name, _text) in enumerate(sections) if name == _MANIFEST_CONTEXT_SECTION
    )
    flexible = tuple(index for index in range(len(sections)) if index not in manifests)
    manifest_total = sum(lengths[index] for index in manifests)
    allocations = [0] * len(sections)

    if manifest_total <= budget:
        for index in manifests:
            allocations[index] = lengths[index]
        remaining = budget - manifest_total

        base_capacities = tuple(
            min(lengths[index], _SAMPLE_SECTION_RESERVE_CHARS) for index in flexible
        )
        reserved = _fair_allocations(base_capacities, remaining)
        for index, allocation in zip(flexible, reserved, strict=True):
            allocations[index] = allocation
        remaining -= sum(reserved)

        if remaining > 0:
            extra_capacities = tuple(lengths[index] - allocations[index] for index in flexible)
            extras = _fair_allocations(extra_capacities, remaining)
            for index, extra in zip(flexible, extras, strict=True):
                allocations[index] += extra
        return tuple(allocations)

    # A manifest larger than the complete caller-supplied budget cannot be
    # retained verbatim.  Apply the same deterministic fair fallback rather
    # than allowing an early prefix to erase every later section.
    return _fair_allocations(lengths, budget)


def _is_child_status(line: str) -> bool:
    return line.startswith(_CHILD_STATUS_PREFIXES)


def _child_blocks(text: str) -> tuple[str, tuple[_ChildBlock, ...]] | None:
    """Split a top-level section into its prefix and ``###`` child blocks."""

    matches = tuple(_CHILD_HEADER_RE.finditer(text))
    if not matches:
        return None

    prefix = text[: matches[0].start()].rstrip("\n")
    blocks: list[_ChildBlock] = []
    for index, match in enumerate(matches):
        end = matches[index + 1].start() if index + 1 < len(matches) else len(text)
        lines = text[match.start() : end].strip("\n").splitlines()
        heading = lines[0]
        remainder = lines[1:]
        while remainder and not remainder[0]:
            remainder.pop(0)

        status_lines: list[str] = []
        while remainder and _is_child_status(remainder[0]):
            status_lines.append(remainder.pop(0))
        while remainder and not remainder[0]:
            remainder.pop(0)

        blocks.append(
            _ChildBlock(
                heading=heading,
                status="\n".join(status_lines),
                body="\n".join(remainder).rstrip("\n"),
            )
        )
    return prefix, tuple(blocks)


def _child_omission_line(count: int) -> str:
    return f"[CHILD BLOCKS OMITTED; count={count}]"


def _render_child_content(
    prefix: str,
    blocks: tuple[_ChildBlock, ...],
    payloads: tuple[str, ...],
    *,
    omitted: int,
) -> str:
    parts = [prefix]
    for block, payload in zip(blocks, payloads, strict=True):
        rendered = block.heading
        if payload:
            rendered += "\n" + payload
        parts.append(rendered)
    if omitted:
        parts.append(_child_omission_line(omitted))
    return "\n\n".join(part for part in parts if part)


def _payload_prefix(payload: str, allocation: int) -> str:
    """Render at most ``allocation`` chars, including the heading separator."""

    if not payload or allocation <= 1:
        return ""
    limit = allocation - 1
    prefix = payload[:limit]
    boundary = prefix.rfind("\n")
    if boundary >= max(0, limit - 400):
        prefix = prefix[:boundary]
    return prefix


def _status_and_body(block: _ChildBlock, allocation: int) -> str:
    body = _payload_prefix(block.body, allocation)
    return "\n".join(part for part in (block.status, body) if part)


def _clip_child_section(
    prefix: str,
    blocks: tuple[_ChildBlock, ...],
    available: int,
) -> str | None:
    """Clip child blocks without allowing early ranges to erase later ones.

    Child headings are the primary structural evidence.  Once every heading
    fits, status lines are protected when possible and the remaining capacity
    is water-filled across child bodies.  If every heading cannot fit, a stable
    leading subset is retained alongside an explicit omitted-block count.
    """

    empty_payloads = tuple("" for _block in blocks)
    visible = blocks
    omitted = 0
    heading_only = _render_child_content(prefix, visible, empty_payloads, omitted=0)

    if len(heading_only) > available:
        selected = None
        for count in range(len(blocks) - 1, -1, -1):
            candidate_blocks = blocks[:count]
            candidate = _render_child_content(
                prefix,
                candidate_blocks,
                tuple("" for _block in candidate_blocks),
                omitted=len(blocks) - count,
            )
            if len(candidate) <= available:
                selected = (candidate_blocks, len(blocks) - count, candidate)
                break
        if selected is None:
            return None
        visible, omitted, heading_only = selected

    status_payloads = tuple(block.status for block in visible)
    with_status = _render_child_content(
        prefix,
        visible,
        status_payloads,
        omitted=omitted,
    )
    if len(with_status) <= available:
        remaining = available - len(with_status)
        body_capacities = tuple(1 + len(block.body) if block.body else 0 for block in visible)
        body_allocations = _fair_allocations(body_capacities, remaining)
        payloads = tuple(
            _status_and_body(block, allocation)
            for block, allocation in zip(visible, body_allocations, strict=True)
        )
        return _render_child_content(prefix, visible, payloads, omitted=omitted)

    # All headings fit but the structural status text does not.  Fairly share
    # what remains across each complete child payload, whose prefix begins with
    # the status line where one exists.
    remaining = available - len(heading_only)
    full_payloads = tuple(
        "\n".join(part for part in (block.status, block.body) if part) for block in visible
    )
    payload_capacities = tuple(1 + len(payload) if payload else 0 for payload in full_payloads)
    allocations = _fair_allocations(payload_capacities, remaining)
    payloads = tuple(
        _payload_prefix(payload, allocation)
        for payload, allocation in zip(full_payloads, allocations, strict=True)
    )
    return _render_child_content(prefix, visible, payloads, omitted=omitted)


def _clip_section(text: str, char_budget: int) -> str:
    if len(text) <= char_budget:
        return text
    if char_budget <= 0:
        return ""

    marker = f"\n[SECTION TRUNCATED; original_chars={len(text)}]\n"
    heading = text.partition("\n")[0] + "\n"
    if char_budget <= len(heading) + len(marker):
        fallback = heading + "[SECTION TRUNCATED]\n"
        return fallback[:char_budget]

    child_parts = _child_blocks(text)
    if child_parts is not None:
        prefix, blocks = child_parts
        clipped = _clip_child_section(prefix, blocks, char_budget - len(marker))
        if clipped is not None:
            return clipped + marker

    prefix_limit = char_budget - len(marker)
    prefix = text[:prefix_limit]
    boundary = prefix.rfind("\n")
    if boundary >= max(len(heading), prefix_limit - 400):
        prefix = prefix[:boundary]
    return prefix + marker


def _clip_context(
    full_text: str,
    char_budget: int,
    *,
    sections: tuple[tuple[str, str], ...] | None = None,
) -> tuple[str, bool]:
    if len(full_text) <= char_budget:
        return full_text, False
    marker = f"\n[CONTEXT TRUNCATED; original_chars={len(full_text)}]\n"
    if char_budget <= len(marker):
        return marker.lstrip("\n")[:char_budget], True

    rendered_sections = _context_sections(full_text) if sections is None else sections
    allocations = _section_allocations(rendered_sections, char_budget - len(marker))
    rendered = "".join(
        _clip_section(section_text, allocation)
        for (_name, section_text), allocation in zip(
            rendered_sections,
            allocations,
            strict=True,
        )
    )
    return rendered + marker, True


def build_context(
    task: TaskSpec,
    char_budget: int = CONTEXT_CHAR_BUDGET,
) -> ContextPack:
    """Build deterministic model context without opening any golden workbook."""

    if isinstance(char_budget, bool) or not isinstance(char_budget, int) or char_budget < 1:
        raise ValueError("char_budget must be a positive integer")
    sections = _render_context_sections(task)
    full_text = "".join(section_text for _name, section_text in sections)
    text, truncated = _clip_context(full_text, char_budget, sections=sections)
    digest = hashlib.sha256(text.encode("utf-8")).hexdigest()
    return ContextPack(
        text=text,
        original_chars=len(full_text),
        truncated=truncated,
        sha256=digest,
    )
