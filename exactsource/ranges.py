"""Parsing and expansion of SpreadsheetBench's occasionally irregular ranges.

The source dataset contains several dialects of sheet-qualified A1 notation.
Besides valid Excel references (``'Sales Data'!A1:B4``), older records may put
the closing quote after the ``!`` or omit one side of the quote.  The parser is
deliberately tolerant of those quote-placement errors while remaining strict
about the actual A1 range.
"""

from __future__ import annotations

import re
from collections.abc import Iterator

from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.utils.cell import range_boundaries

from exactsource.contracts import QualifiedRange


class RangeSyntaxError(ValueError):
    """Raised when a SpreadsheetBench range cannot be interpreted safely."""


class RangeTooLargeError(ValueError):
    """Raised when exact expansion exceeds a caller-supplied safety limit."""


_CELL = r"\$?[A-Za-z]{1,3}\$?\d+"
_COLUMN = r"\$?[A-Za-z]{1,3}"
_ROW = r"\$?\d+"
_RANGE_TOKEN = (
    rf"(?:{_CELL}\s*:\s*(?:{_CELL}|{_ROW})|{_CELL}|{_COLUMN}\s*:\s*{_COLUMN}|{_ROW}\s*:\s*{_ROW})"
)
_REFERENCE_RE = re.compile(
    rf"(?<![A-Za-z0-9_$])(?P<cells>{_RANGE_TOKEN})(?![A-Za-z0-9_])",
    flags=re.IGNORECASE,
)
_CELL_RE = re.compile(r"^\$?(?P<column>[A-Za-z]{1,3})\$?(?P<row>\d+)$")
_COLUMN_RE = re.compile(r"^\$?(?P<column>[A-Za-z]{1,3})$")
_ROW_RE = re.compile(r"^\$?(?P<row>\d+)$")
_SAFE_UNQUOTED_SHEET_RE = re.compile(r"^[A-Za-z_][A-Za-z0-9_.]*$")
MAX_EXCEL_ROW = 1_048_576
MAX_EXCEL_COLUMN = 16_384


def _normalise_quotes(value: str) -> str:
    return (
        value.replace("\u00a0", " ")
        .replace("\u2018", "'")
        .replace("\u2019", "'")
        .replace("\u201c", '"')
        .replace("\u201d", '"')
    )


def _clean_sheet_name(raw: str) -> str:
    """Remove dataset delimiter debris without changing interior punctuation."""

    value = _normalise_quotes(raw).strip()
    previous = None
    while value != previous:
        previous = value
        value = value.strip(" \t\r\n,;")
        if len(value) >= 2 and value[0] == value[-1] and value[0] in "'\"":
            value = value[1:-1].strip()
        else:
            value = value.strip("'\"").strip()
    value = value.replace("''", "'")
    if not value:
        raise RangeSyntaxError(f"empty sheet name in {raw!r}")
    return value


def _looks_like_sheet_label(text: str, end: int) -> bool:
    """Reject an A1-looking sheet name such as ``'A1'!B2``."""

    return re.match(r"\s*['\"]?\s*!", text[end:]) is not None


def normalise_a1_range(value: str) -> str:
    """Return canonical, absolute-marker-free A1 notation.

    SpreadsheetBench includes legacy forms such as ``BD2:308``.  Consistent
    with its official evaluator, a numeric range end inherits the start
    column, producing ``BD2:BD308``.
    """

    raw = _normalise_quotes(value).strip().strip("'\"")
    compact = re.sub(r"\s+", "", raw).replace("$", "")
    if not compact:
        raise RangeSyntaxError("range is empty")

    if ":" not in compact:
        match = _CELL_RE.fullmatch(compact)
        if not match:
            raise RangeSyntaxError(f"not an A1 cell reference: {value!r}")
        row = int(match.group("row"))
        column = match.group("column").upper()
        if row < 1 or row > MAX_EXCEL_ROW or column_index_from_string(column) > MAX_EXCEL_COLUMN:
            raise RangeSyntaxError(f"row must be positive: {value!r}")
        return f"{column}{row}"

    start_raw, end_raw = compact.split(":", 1)
    start_cell = _CELL_RE.fullmatch(start_raw)
    end_cell = _CELL_RE.fullmatch(end_raw)
    start_col = _COLUMN_RE.fullmatch(start_raw)
    end_col = _COLUMN_RE.fullmatch(end_raw)
    start_row = _ROW_RE.fullmatch(start_raw)
    end_row = _ROW_RE.fullmatch(end_raw)

    if start_cell and end_row:
        end_number = int(end_row.group("row"))
        if end_number < 1:
            raise RangeSyntaxError(f"row must be positive: {value!r}")
        end_raw = f"{start_cell.group('column')}{end_number}"
        end_cell = _CELL_RE.fullmatch(end_raw)

    if start_cell and end_cell:
        start_number = int(start_cell.group("row"))
        end_number = int(end_cell.group("row"))
        start_column = start_cell.group("column").upper()
        end_column = end_cell.group("column").upper()
        if (
            start_number < 1
            or end_number < start_number
            or end_number > MAX_EXCEL_ROW
            or column_index_from_string(start_column) > MAX_EXCEL_COLUMN
            or column_index_from_string(end_column) > MAX_EXCEL_COLUMN
            or column_index_from_string(end_column) < column_index_from_string(start_column)
        ):
            raise RangeSyntaxError(
                f"range must fit Excel and run top-left to bottom-right: {value!r}"
            )
        return f"{start_column}{start_number}:{end_column}{end_number}"

    if start_col and end_col:
        start_column = start_col.group("column").upper()
        end_column = end_col.group("column").upper()
        if (
            column_index_from_string(start_column) > MAX_EXCEL_COLUMN
            or column_index_from_string(end_column) > MAX_EXCEL_COLUMN
            or column_index_from_string(end_column) < column_index_from_string(start_column)
        ):
            raise RangeSyntaxError(f"column range must fit Excel and run left to right: {value!r}")
        return f"{start_column}:{end_column}"

    if start_row and end_row:
        start_number = int(start_row.group("row"))
        end_number = int(end_row.group("row"))
        if start_number < 1 or end_number < start_number or end_number > MAX_EXCEL_ROW:
            raise RangeSyntaxError(f"row range must fit Excel and run top to bottom: {value!r}")
        return f"{start_number}:{end_number}"

    raise RangeSyntaxError(f"not a supported A1 range: {value!r}")


def sheet_candidates(value: str | None) -> tuple[str, ...]:
    """Parse the dataset's answer-sheet field into ordered candidates.

    A single properly quoted name may contain commas.  Multiple names appear
    either as ``A,B`` or ``'A','B'``.  The first candidate is the evaluator's
    effective default for an unqualified answer range.
    """

    if value is None or not str(value).strip():
        return ()
    text = _normalise_quotes(str(value)).strip()
    if len(text) >= 2 and text[0] == text[-1] == "'" and "','" not in text:
        return (_clean_sheet_name(text),)

    chunks = re.split(r"\s*['\"]?\s*,\s*['\"]?\s*", text)
    result: list[str] = []
    for chunk in chunks:
        if not chunk.strip(" \t\r\n'\""):
            continue
        cleaned = _clean_sheet_name(chunk)
        if cleaned not in result:
            result.append(cleaned)
    return tuple(result)


def parse_qualified_ranges(
    value: str,
    default_sheet: str | None = None,
) -> tuple[QualifiedRange, ...]:
    """Extract every qualified or unqualified A1 range from ``value``.

    Coordinates are located first, then the text immediately before each
    coordinate is interpreted as its sheet qualifier.  This makes the parser
    robust to the malformed quote placement found in the benchmark while also
    preserving valid names containing commas, such as ``'b2b, sez, de'``.
    """

    if not isinstance(value, str) or not value.strip():
        raise RangeSyntaxError("range expression must be a non-empty string")
    text = _normalise_quotes(value)
    default = _clean_sheet_name(default_sheet) if default_sheet and default_sheet.strip() else None

    parsed: list[QualifiedRange] = []
    previous_end = 0
    for match in _REFERENCE_RE.finditer(text):
        if _looks_like_sheet_label(text, match.end()):
            continue

        prefix = text[previous_end : match.start()]
        sheet = default
        if "!" in prefix:
            sheet = _clean_sheet_name(prefix.rsplit("!", 1)[0])
        if sheet is None:
            raise RangeSyntaxError(f"range {match.group('cells')!r} has no sheet")

        cells = normalise_a1_range(match.group("cells"))
        parsed.append(QualifiedRange(sheet=sheet, cells=cells))
        previous_end = match.end()

    if not parsed:
        raise RangeSyntaxError(f"no A1 range found in {value!r}")
    return tuple(parsed)


def parse_answer_ranges(
    answer_position: str, answer_sheet: str | None
) -> tuple[QualifiedRange, ...]:
    """Parse a task's graded ranges using its first answer sheet as default."""

    candidates = sheet_candidates(answer_sheet)
    default = candidates[0] if candidates else None
    return parse_qualified_ranges(answer_position, default_sheet=default)


def range_bounds(
    cells: str,
    *,
    max_row: int | None = None,
    max_column: int | None = None,
) -> tuple[int, int, int, int]:
    """Resolve a canonical range to inclusive numeric bounds.

    Whole-column and whole-row references require the corresponding worksheet
    dimension.  This prevents accidental expansion to Excel's full grid.
    """

    canonical = normalise_a1_range(cells)
    min_col, min_row, final_col, final_row = range_boundaries(canonical)
    if min_row is None or final_row is None:
        if max_row is None:
            raise RangeSyntaxError(f"whole-column range {cells!r} requires max_row")
        min_row, final_row = 1, max(1, int(max_row))
    if min_col is None or final_col is None:
        if max_column is None:
            raise RangeSyntaxError(f"whole-row range {cells!r} requires max_column")
        min_col, final_col = 1, max(1, int(max_column))
    return int(min_col), int(min_row), int(final_col), int(final_row)


def range_cell_count(
    cells: str,
    *,
    max_row: int | None = None,
    max_column: int | None = None,
) -> int:
    min_col, min_row, final_col, final_row = range_bounds(
        cells,
        max_row=max_row,
        max_column=max_column,
    )
    return (final_col - min_col + 1) * (final_row - min_row + 1)


def iter_range_coordinates(
    cells: str,
    *,
    max_row: int | None = None,
    max_column: int | None = None,
    max_cells: int | None = None,
) -> Iterator[str]:
    """Yield coordinates in the same row-major order as the evaluator."""

    min_col, min_row, final_col, final_row = range_bounds(
        cells,
        max_row=max_row,
        max_column=max_column,
    )
    count = (final_col - min_col + 1) * (final_row - min_row + 1)
    if max_cells is not None and count > max_cells:
        raise RangeTooLargeError(f"{cells} contains {count} cells; limit is {max_cells}")
    for row in range(min_row, final_row + 1):
        for column in range(min_col, final_col + 1):
            yield f"{get_column_letter(column)}{row}"


def quote_sheet_name(sheet: str) -> str:
    """Return an Excel-safe sheet token for display and prompt construction."""

    if _SAFE_UNQUOTED_SHEET_RE.fullmatch(sheet):
        return sheet
    return "'" + sheet.replace("'", "''") + "'"


def format_qualified_range(value: QualifiedRange) -> str:
    return f"{quote_sheet_name(value.sheet)}!{value.cells}"
