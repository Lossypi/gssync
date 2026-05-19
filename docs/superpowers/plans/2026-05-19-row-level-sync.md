# Row-Level Sync Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add 4 MCP tools (`get_rows`, `update_rows`, `pull_rows`, `push_rows`) for reading and writing individual rows in Google Sheets without touching an entire sheet.

**Architecture:** New `gssync/rows.py` holds all filter/read/write logic. `gssync/mcp_server.py` gets 4 new tools that import from `rows.py`. Existing tools and `sync.py` are not modified.

**Tech Stack:** Python 3.10+, gspread, existing gssync modules (storage, sheets, config)

---

## File Map

| File | Action | Responsibility |
|------|--------|----------------|
| `gssync/rows.py` | Create | Filter parsing, GS read/write, local file read |
| `gssync/mcp_server.py` | Modify | Add 4 new MCP tools + `_format_rows_table` helper |
| `tests/test_rows.py` | Create | Unit tests for all rows.py functions |
| `tests/test_mcp_server.py` | Modify | Tests for 4 new MCP tools |

---

### Task 1: rows.py — filter utilities and read/write functions

**Files:**
- Create: `gssync/rows.py`
- Create: `tests/test_rows.py`

- [ ] **Step 1: Write failing tests for filter utilities**

Create `tests/test_rows.py`:

```python
import pytest
from pathlib import Path
from unittest.mock import MagicMock, patch

from gssync.rows import (
    _col_index_to_letter,
    _col_letter_to_index,
    filter_rows_by_column,
    parse_cell_range,
    parse_row_numbers,
    read_rows_from_local,
    read_rows_from_sheet,
    resolve_filter,
    rows_data_to_lists,
    write_rows_to_sheet,
)


# ── parse_row_numbers ────────────────────────────────────────────────────────

def test_parse_row_numbers_single():
    assert parse_row_numbers("5") == [5]

def test_parse_row_numbers_list():
    assert parse_row_numbers("2,5,10") == [2, 5, 10]

def test_parse_row_numbers_range():
    assert parse_row_numbers("2:5") == [2, 3, 4, 5]

def test_parse_row_numbers_single_item_list():
    assert parse_row_numbers("3,") == [3]


# ── _col_letter_to_index / _col_index_to_letter ──────────────────────────────

def test_col_letter_to_index():
    assert _col_letter_to_index("A") == 1
    assert _col_letter_to_index("Z") == 26
    assert _col_letter_to_index("AA") == 27

def test_col_index_to_letter():
    assert _col_index_to_letter(1) == "A"
    assert _col_index_to_letter(26) == "Z"
    assert _col_index_to_letter(27) == "AA"

def test_col_roundtrip():
    for i in range(1, 30):
        assert _col_letter_to_index(_col_index_to_letter(i)) == i


# ── parse_cell_range ─────────────────────────────────────────────────────────

def test_parse_cell_range_simple():
    assert parse_cell_range("A2:D10") == (2, 10, 1, 4)

def test_parse_cell_range_lowercase():
    assert parse_cell_range("a2:d10") == (2, 10, 1, 4)

def test_parse_cell_range_invalid():
    with pytest.raises(ValueError, match="Invalid cell range"):
        parse_cell_range("bad")


# ── filter_rows_by_column ────────────────────────────────────────────────────

def test_filter_rows_by_column_matches():
    header = ["name", "status"]
    rows = [["Ivan", "active"], ["Maria", "done"], ["Petr", "active"]]
    assert filter_rows_by_column(rows, header, "status", "active") == [1, 3]

def test_filter_rows_by_column_no_match():
    header = ["name", "status"]
    rows = [["Ivan", "active"]]
    assert filter_rows_by_column(rows, header, "status", "done") == []

def test_filter_rows_by_column_missing_column():
    with pytest.raises(ValueError, match="Column 'missing'"):
        filter_rows_by_column([], ["name"], "missing", "x")


# ── resolve_filter ───────────────────────────────────────────────────────────

def test_resolve_filter_row_numbers():
    assert resolve_filter([], row_numbers="2:4") == [2, 3, 4]

def test_resolve_filter_column():
    all_rows = [["name", "status"], ["Ivan", "active"], ["Maria", "done"]]
    assert resolve_filter(all_rows, filter_column="status", filter_value="active") == [1]

def test_resolve_filter_cell_range():
    assert resolve_filter([], cell_range="A3:D5") == [2, 3, 4]

def test_resolve_filter_no_filter():
    with pytest.raises(ValueError, match="Provide exactly one filter"):
        resolve_filter([])

def test_resolve_filter_multiple_filters():
    with pytest.raises(ValueError, match="Provide exactly one filter"):
        resolve_filter([], row_numbers="1", cell_range="A2:B3")


# ── rows_data_to_lists ───────────────────────────────────────────────────────

def test_rows_data_to_lists_basic():
    header = ["name", "status"]
    result = rows_data_to_lists('[{"name": "Ivan", "status": "done"}]', header)
    assert result == [["Ivan", "done"]]

def test_rows_data_to_lists_missing_key():
    header = ["name", "status"]
    result = rows_data_to_lists('[{"name": "Ivan"}]', header)
    assert result == [["Ivan", ""]]

def test_rows_data_to_lists_invalid_json():
    with pytest.raises(ValueError, match="valid JSON"):
        rows_data_to_lists("not json", ["name"])

def test_rows_data_to_lists_not_list():
    with pytest.raises(ValueError, match="JSON array"):
        rows_data_to_lists('{"name": "Ivan"}', ["name"])


# ── read_rows_from_sheet ─────────────────────────────────────────────────────

def test_read_rows_from_sheet_by_indices():
    mock_ws = MagicMock()
    mock_ws.get_all_values.return_value = [
        ["name", "status"],
        ["Ivan", "active"],
        ["Maria", "done"],
        ["Petr", "active"],
    ]
    mock_ss = MagicMock()
    mock_ss.worksheet.return_value = mock_ws
    header, rows = read_rows_from_sheet(mock_ss, "Sheet1", [1, 3])
    assert header == ["name", "status"]
    assert rows == [["Ivan", "active"], ["Petr", "active"]]

def test_read_rows_from_sheet_empty():
    mock_ws = MagicMock()
    mock_ws.get_all_values.return_value = []
    mock_ss = MagicMock()
    mock_ss.worksheet.return_value = mock_ws
    header, rows = read_rows_from_sheet(mock_ss, "Sheet1", [1])
    assert header == []
    assert rows == []

def test_read_rows_from_sheet_out_of_bounds_skipped():
    mock_ws = MagicMock()
    mock_ws.get_all_values.return_value = [["name"], ["Ivan"]]
    mock_ss = MagicMock()
    mock_ss.worksheet.return_value = mock_ws
    header, rows = read_rows_from_sheet(mock_ss, "Sheet1", [1, 99])
    assert rows == [["Ivan"]]


# ── write_rows_to_sheet ──────────────────────────────────────────────────────

def test_write_rows_to_sheet_calls_update():
    mock_ws = MagicMock()
    mock_ss = MagicMock()
    mock_ss.worksheet.return_value = mock_ws
    count = write_rows_to_sheet(mock_ss, "Sheet1", [2], [["Ivan", "done"]])
    assert count == 1
    mock_ws.update.assert_called_once_with(
        "A3:B3", [["Ivan", "done"]], value_input_option="USER_ENTERED"
    )


# ── read_rows_from_local ─────────────────────────────────────────────────────

def test_read_rows_from_local_basic(tmp_path):
    from gssync.storage import write_local
    xlsx = tmp_path / "data.xlsx"
    write_local(xlsx, "xlsx", {"Sheet1": [["name", "status"], ["Ivan", "active"], ["Maria", "done"]]})
    header, rows = read_rows_from_local(xlsx, "xlsx", "Sheet1", [2])
    assert header == ["name", "status"]
    assert rows == [["Maria", "done"]]

def test_read_rows_from_local_missing_sheet(tmp_path):
    from gssync.storage import write_local
    xlsx = tmp_path / "data.xlsx"
    write_local(xlsx, "xlsx", {"Sheet1": [["name"]]})
    with pytest.raises(ValueError, match="Sheet 'Missing'"):
        read_rows_from_local(xlsx, "xlsx", "Missing", [1])
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `py -m pytest tests/test_rows.py -v`

Expected: `ImportError` — `gssync.rows` does not exist yet

- [ ] **Step 3: Create gssync/rows.py**

Create `gssync/rows.py`:

```python
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
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `py -m pytest tests/test_rows.py -v`

Expected: all tests PASS

- [ ] **Step 5: Commit**

```bash
git add gssync/rows.py tests/test_rows.py
git commit -m "feat: add rows.py with filter utilities and GS/local read-write"
```

---

### Task 2: MCP tools — get_rows, update_rows, pull_rows, push_rows

**Files:**
- Modify: `gssync/mcp_server.py`
- Modify: `tests/test_mcp_server.py`

- [ ] **Step 1: Update imports and add failing tests in tests/test_mcp_server.py**

Replace the import block at the top of `tests/test_mcp_server.py` with:

```python
import pytest
from pathlib import Path
from unittest.mock import MagicMock, patch

from gssync.config import Config
from gssync.mcp_server import (
    get_config,
    get_rows,
    list_google_sheets,
    list_local_sheets,
    pull_all,
    pull_rows,
    pull_sheet,
    push_all,
    push_rows,
    push_sheet,
    set_config,
    update_rows,
)
```

Then append these tests to the end of `tests/test_mcp_server.py`:

```python
# ── get_rows ──────────────────────────────────────────────────────────────────

def test_get_rows_by_row_numbers():
    mock_ws = MagicMock()
    mock_ws.get_all_values.return_value = [
        ["name", "status"],
        ["Ivan", "active"],
        ["Maria", "done"],
    ]
    mock_ss = MagicMock()
    mock_ss.worksheet.return_value = mock_ws
    with patch("gssync.mcp_server.get_client"), \
         patch("gssync.mcp_server.open_spreadsheet", return_value=mock_ss):
        result = get_rows(
            spreadsheet_url="https://docs.google.com/spreadsheets/d/abc",
            sheet_name="Sheet1",
            row_numbers="1",
        )
    assert "Ivan" in result
    assert "active" in result


def test_get_rows_no_filter_raises():
    mock_ws = MagicMock()
    mock_ws.get_all_values.return_value = [["name"], ["Ivan"]]
    mock_ss = MagicMock()
    mock_ss.worksheet.return_value = mock_ws
    with patch("gssync.mcp_server.get_client"), \
         patch("gssync.mcp_server.open_spreadsheet", return_value=mock_ss):
        import pytest
        with pytest.raises(ValueError, match="Provide exactly one filter"):
            get_rows(
                spreadsheet_url="https://docs.google.com/spreadsheets/d/abc",
                sheet_name="Sheet1",
            )


# ── update_rows ───────────────────────────────────────────────────────────────

def test_update_rows_by_row_numbers():
    mock_ws = MagicMock()
    mock_ws.get_all_values.return_value = [
        ["name", "status"],
        ["Ivan", "active"],
    ]
    mock_ss = MagicMock()
    mock_ss.worksheet.return_value = mock_ws
    with patch("gssync.mcp_server.get_client"), \
         patch("gssync.mcp_server.open_spreadsheet", return_value=mock_ss):
        result = update_rows(
            spreadsheet_url="https://docs.google.com/spreadsheets/d/abc",
            sheet_name="Sheet1",
            rows_data='[{"name": "Ivan", "status": "done"}]',
            row_numbers="1",
        )
    mock_ws.update.assert_called_once_with(
        "A2:B2", [["Ivan", "done"]], value_input_option="USER_ENTERED"
    )
    assert "1" in result


def test_update_rows_broadcast_filter_column():
    mock_ws = MagicMock()
    mock_ws.get_all_values.return_value = [
        ["name", "status"],
        ["Ivan", "active"],
        ["Maria", "active"],
    ]
    mock_ss = MagicMock()
    mock_ss.worksheet.return_value = mock_ws
    with patch("gssync.mcp_server.get_client"), \
         patch("gssync.mcp_server.open_spreadsheet", return_value=mock_ss):
        result = update_rows(
            spreadsheet_url="https://docs.google.com/spreadsheets/d/abc",
            sheet_name="Sheet1",
            rows_data='[{"name": "Ivan", "status": "done"}]',
            filter_column="status",
            filter_value="active",
        )
    assert mock_ws.update.call_count == 2
    assert "2" in result


# ── pull_rows ─────────────────────────────────────────────────────────────────

def test_pull_rows_writes_filtered_rows_to_file(tmp_path):
    mock_ws = MagicMock()
    mock_ws.get_all_values.return_value = [
        ["name", "status"],
        ["Ivan", "active"],
        ["Maria", "done"],
    ]
    mock_ss = MagicMock()
    mock_ss.worksheet.return_value = mock_ws
    xlsx = tmp_path / "data.xlsx"
    with patch("gssync.mcp_server.get_client"), \
         patch("gssync.mcp_server.open_spreadsheet", return_value=mock_ss):
        result = pull_rows(
            spreadsheet_url="https://docs.google.com/spreadsheets/d/abc",
            sheet_name="Sheet1",
            file_path=str(xlsx),
            row_numbers="1",
        )
    from gssync.storage import read_local
    data = read_local(xlsx, "xlsx")
    assert data["Sheet1"] == [["name", "status"], ["Ivan", "active"]]
    assert "1" in result


# ── push_rows ─────────────────────────────────────────────────────────────────

def test_push_rows_from_local_file(tmp_path):
    from gssync.storage import write_local
    xlsx = tmp_path / "data.xlsx"
    write_local(xlsx, "xlsx", {
        "Sheet1": [["name", "status"], ["Ivan", "active"], ["Maria", "done"]]
    })
    mock_ws = MagicMock()
    mock_ss = MagicMock()
    mock_ss.worksheet.return_value = mock_ws
    with patch("gssync.mcp_server.get_client"), \
         patch("gssync.mcp_server.open_spreadsheet", return_value=mock_ss):
        result = push_rows(
            spreadsheet_url="https://docs.google.com/spreadsheets/d/abc",
            sheet_name="Sheet1",
            file_path=str(xlsx),
            row_numbers="2",
        )
    mock_ws.update.assert_called_once_with(
        "A3:B3", [["Maria", "done"]], value_input_option="USER_ENTERED"
    )
    assert "1" in result
```

- [ ] **Step 2: Run tests to verify new tests fail**

Run: `py -m pytest tests/test_mcp_server.py -v`

Expected: existing 10 tests PASS, new tests FAIL with `ImportError` for `get_rows`, `update_rows`, `pull_rows`, `push_rows`

- [ ] **Step 3: Add 4 new tools to gssync/mcp_server.py**

Add this import at the top of `gssync/mcp_server.py`, after the existing imports:

```python
from .rows import (
    read_range_from_sheet,
    read_rows_from_local,
    read_rows_from_sheet,
    resolve_filter,
    rows_data_to_lists,
    write_range_to_sheet,
    write_rows_to_sheet,
)
```

Add this private helper function in `gssync/mcp_server.py` before `if __name__ == "__main__":`:

```python
def _format_rows_table(header: list, rows: list) -> str:
    all_data = [header] + rows
    widths = [
        max(len(str(row[i]) if i < len(row) else "") for row in all_data)
        for i in range(len(header))
    ]
    sep = "-+-".join("-" * w for w in widths)
    def fmt(row):
        return " | ".join(
            str(row[i] if i < len(row) else "").ljust(widths[i])
            for i in range(len(header))
        )
    lines = [fmt(header), sep] + [fmt(r) for r in rows]
    return "\n".join(lines)
```

Add these 4 tool functions in `gssync/mcp_server.py` before `_format_rows_table`:

```python
@mcp.tool()
def get_rows(
    spreadsheet_url: str,
    sheet_name: str,
    row_numbers: str = "",
    filter_column: str = "",
    filter_value: str = "",
    cell_range: str = "",
) -> str:
    client = get_client()
    spreadsheet = open_spreadsheet(client, spreadsheet_url)
    ws = spreadsheet.worksheet(sheet_name)
    if cell_range and not row_numbers and not filter_column:
        header, rows = read_range_from_sheet(spreadsheet, sheet_name, cell_range)
    else:
        all_rows = ws.get_all_values(value_render_option="FORMULA")
        indices = resolve_filter(all_rows, row_numbers, filter_column, filter_value, cell_range)
        header, rows = read_rows_from_sheet(spreadsheet, sheet_name, indices)
    if not rows:
        return f"No rows found in '{sheet_name}'"
    return f"Sheet: {sheet_name} | {len(rows)} row(s)\n{_format_rows_table(header, rows)}"


@mcp.tool()
def update_rows(
    spreadsheet_url: str,
    sheet_name: str,
    rows_data: str,
    row_numbers: str = "",
    filter_column: str = "",
    filter_value: str = "",
    cell_range: str = "",
) -> str:
    import json as _json
    client = get_client()
    spreadsheet = open_spreadsheet(client, spreadsheet_url)
    ws = spreadsheet.worksheet(sheet_name)
    if cell_range and not row_numbers and not filter_column:
        # Direct range write — rows_data must be list of lists
        data = _json.loads(rows_data)
        count = write_range_to_sheet(spreadsheet, sheet_name, cell_range, data)
        return f"Updated {count} row(s) in '{sheet_name}' at {cell_range}"
    all_rows = ws.get_all_values(value_render_option="FORMULA")
    header = all_rows[0] if all_rows else []
    indices = resolve_filter(all_rows, row_numbers, filter_column, filter_value)
    new_rows = rows_data_to_lists(rows_data, header)
    if filter_column:
        # Broadcast: apply first entry to all matching rows
        broadcast = new_rows[0] if new_rows else []
        write_data = [broadcast] * len(indices)
    else:
        write_data = new_rows
    count = write_rows_to_sheet(spreadsheet, sheet_name, indices, write_data)
    return f"Updated {count} row(s) in '{sheet_name}'"


@mcp.tool()
def pull_rows(
    spreadsheet_url: str,
    sheet_name: str,
    file_path: str,
    file_format: str = "xlsx",
    row_numbers: str = "",
    filter_column: str = "",
    filter_value: str = "",
    cell_range: str = "",
) -> str:
    from .storage import read_local, write_local
    client = get_client()
    spreadsheet = open_spreadsheet(client, spreadsheet_url)
    ws = spreadsheet.worksheet(sheet_name)
    all_rows = ws.get_all_values(value_render_option="FORMULA")
    indices = resolve_filter(all_rows, row_numbers, filter_column, filter_value, cell_range)
    header, rows = read_rows_from_sheet(spreadsheet, sheet_name, indices)
    path = Path(file_path)
    exists = (path.exists() and path.is_dir()) if file_format == "csv" else path.exists()
    existing = read_local(path, file_format) if exists else {}
    existing[sheet_name] = [header] + rows
    write_local(path, file_format, existing)
    return f"Pulled {len(rows)} row(s) from '{sheet_name}' → {file_path}"


@mcp.tool()
def push_rows(
    spreadsheet_url: str,
    sheet_name: str,
    file_path: str,
    file_format: str = "xlsx",
    row_numbers: str = "",
    filter_column: str = "",
    filter_value: str = "",
    cell_range: str = "",
) -> str:
    from .storage import read_local
    client = get_client()
    spreadsheet = open_spreadsheet(client, spreadsheet_url)
    path = Path(file_path)
    local_data = read_local(path, file_format)
    all_rows = local_data.get(sheet_name, [])
    indices = resolve_filter(all_rows, row_numbers, filter_column, filter_value, cell_range)
    _, rows = read_rows_from_local(path, file_format, sheet_name, indices)
    count = write_rows_to_sheet(spreadsheet, sheet_name, indices, rows)
    return f"Pushed {count} row(s) from {file_path} → '{sheet_name}'"
```

- [ ] **Step 4: Run all MCP tests**

Run: `py -m pytest tests/test_mcp_server.py -v`

Expected: all tests PASS (10 existing + new row tool tests)

- [ ] **Step 5: Run full test suite**

Run: `py -m pytest -v`

Expected: all tests PASS

- [ ] **Step 6: Smoke test**

Run: `py -c "from gssync.mcp_server import mcp; print('OK')"`

Expected: `OK`

- [ ] **Step 7: Commit**

```bash
git add gssync/mcp_server.py tests/test_mcp_server.py
git commit -m "feat: add get_rows, update_rows, pull_rows, push_rows MCP tools"
```
