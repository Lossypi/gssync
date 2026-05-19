import json
import re
from pathlib import Path
from typing import List, Tuple

import gspread

from .storage import read_local


# ── Column helpers ────────────────────────────────────────────────────────────

def _col_letter_to_index(letters: str) -> int:
    """Convert column letter(s) to 1-based index. A→1, Z→26, AA→27."""
    result = 0
    for ch in letters.upper():
        result = result * 26 + (ord(ch) - ord("A") + 1)
    return result


def _col_index_to_letter(idx: int) -> str:
    """Convert 1-based column index to letter(s). 1→A, 26→Z, 27→AA."""
    result = ""
    while idx > 0:
        idx, remainder = divmod(idx - 1, 26)
        result = chr(65 + remainder) + result
    return result


# ── Filter utilities ──────────────────────────────────────────────────────────

def parse_row_numbers(spec: str) -> List[int]:
    """Parse "5", "2,5,10", or "2:10" into list of 1-based data row indices."""
    spec = spec.strip()
    if ":" in spec:
        start, end = spec.split(":", 1)
        return list(range(int(start), int(end) + 1))
    return [int(x.strip()) for x in spec.split(",") if x.strip()]


def parse_cell_range(cell_range: str) -> Tuple[int, int, int, int]:
    """Parse "A2:D10" → (row_start, row_end, col_start, col_end), 1-based."""
    m = re.match(r"^([A-Z]+)(\d+):([A-Z]+)(\d+)$", cell_range.upper().strip())
    if not m:
        raise ValueError(f"Invalid cell range: '{cell_range}'")
    return (
        int(m.group(2)),
        int(m.group(4)),
        _col_letter_to_index(m.group(1)),
        _col_letter_to_index(m.group(3)),
    )


def filter_rows_by_column(
    data_rows: List[List], header: List, column: str, value: str
) -> List[int]:
    """Return 1-based data row indices where header[column] == value."""
    if column not in header:
        raise ValueError(f"Column '{column}' not found in sheet")
    col_idx = header.index(column)
    return [
        i + 1
        for i, row in enumerate(data_rows)
        if len(row) > col_idx and str(row[col_idx]) == value
    ]


def resolve_filter(
    all_rows: List[List],
    row_numbers: str = "",
    filter_column: str = "",
    filter_value: str = "",
    cell_range: str = "",
) -> List[int]:
    """Resolve filter parameters to 1-based data row indices.
    Exactly one of row_numbers, filter_column, or cell_range must be provided.
    For cell_range, column constraints are ignored — only the row range is used."""
    provided = sum([bool(row_numbers), bool(filter_column), bool(cell_range)])
    if provided == 0:
        raise ValueError(
            "Provide exactly one filter: row_numbers, filter_column/filter_value, or cell_range"
        )
    if provided > 1:
        raise ValueError(
            "Provide exactly one filter: row_numbers, filter_column/filter_value, or cell_range"
        )
    if row_numbers:
        return parse_row_numbers(row_numbers)
    if filter_column:
        header = all_rows[0] if all_rows else []
        data = all_rows[1:] if len(all_rows) > 1 else []
        return filter_rows_by_column(data, header, filter_column, filter_value)
    # cell_range: convert sheet rows to 1-based data row indices (subtract 1 for header)
    row_start, row_end, _, _ = parse_cell_range(cell_range)
    if row_start < 2:
        raise ValueError(
            "cell_range must start at row 2 or later (row 1 is the header)"
        )
    return list(range(row_start - 1, row_end))


def rows_data_to_lists(rows_data_json: str, header: List) -> List[List]:
    """Convert JSON list-of-dicts to list-of-lists ordered by header.
    Missing keys become empty string."""
    try:
        data = json.loads(rows_data_json)
    except json.JSONDecodeError as e:
        raise ValueError(f"rows_data must be valid JSON: {e}")
    if not isinstance(data, list):
        raise ValueError("rows_data must be a JSON array of objects")
    return [[row_dict.get(col, "") for col in header] for row_dict in data]


# ── GS read ───────────────────────────────────────────────────────────────────

def read_rows_from_sheet(
    spreadsheet: gspread.Spreadsheet, sheet_name: str, indices: List[int]
) -> Tuple[List, List[List]]:
    """Read specific data rows from GS.
    indices are 1-based data row numbers (row 1 = first row after header).
    Out-of-bounds indices are silently skipped.
    Returns (header, selected_rows)."""
    ws = spreadsheet.worksheet(sheet_name)
    all_rows = ws.get_all_values(value_render_option="FORMULA")
    if not all_rows:
        return [], []
    header = all_rows[0]
    data_rows = all_rows[1:]
    selected = [data_rows[i - 1] for i in indices if 1 <= i <= len(data_rows)]
    return header, selected


def read_range_from_sheet(
    spreadsheet: gspread.Spreadsheet, sheet_name: str, cell_range: str
) -> Tuple[List, List[List]]:
    """Read a cell range from GS. Also fetches row 1 as header for context.
    Returns (header, range_rows)."""
    ws = spreadsheet.worksheet(sheet_name)
    header = ws.row_values(1)
    range_rows = ws.get(cell_range, value_render_option="FORMULA")
    return header, range_rows or []


# ── GS write ──────────────────────────────────────────────────────────────────

def write_rows_to_sheet(
    spreadsheet: gspread.Spreadsheet, sheet_name: str, indices: List[int], rows: List[List]
) -> int:
    """Write rows to specific data positions in GS.
    indices are 1-based data row numbers (row 1 = first row after header, i.e. sheet row 2).
    Returns count of rows written."""
    if len(indices) != len(rows):
        raise ValueError(
            f"indices and rows must have the same length (got {len(indices)} and {len(rows)})"
        )
    ws = spreadsheet.worksheet(sheet_name)
    count = 0
    for idx, row in zip(indices, rows):
        sheet_row = idx + 1  # +1 because row 1 is the header
        end_col = _col_index_to_letter(max(len(row), 1))
        ws.update(
            f"A{sheet_row}:{end_col}{sheet_row}",
            [row],
            value_input_option="USER_ENTERED",
        )
        count += 1
    return count


def write_range_to_sheet(
    spreadsheet: gspread.Spreadsheet, sheet_name: str, cell_range: str, rows: List[List]
) -> int:
    """Write a 2D list directly to a cell range in GS. Returns count of rows written."""
    ws = spreadsheet.worksheet(sheet_name)
    ws.update(cell_range, rows, value_input_option="USER_ENTERED")
    return len(rows)


# ── Local file read ───────────────────────────────────────────────────────────

def read_rows_from_local(
    path: Path, fmt: str, sheet_name: str, indices: List[int]
) -> Tuple[List, List[List]]:
    """Read specific data rows from a local file.
    indices are 1-based data row numbers (row 1 = first row after header).
    Returns (header, selected_rows)."""
    data = read_local(path, fmt)
    if sheet_name not in data:
        raise ValueError(f"Sheet '{sheet_name}' not found in {path}")
    all_rows = data[sheet_name]
    if not all_rows:
        return [], []
    header = all_rows[0]
    data_rows = all_rows[1:]
    selected = [data_rows[i - 1] for i in indices if 1 <= i <= len(data_rows)]
    return header, selected
