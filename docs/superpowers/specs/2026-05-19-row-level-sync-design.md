# Row-Level Sync Design

## Goal

Add row-level read/write tools to GSSync so Claude can read and update individual rows in Google Sheets without touching an entire sheet.

## Architecture

New file `gssync/rows.py` contains all row-level logic (filtering, reading, writing). `gssync/mcp_server.py` gets 4 new tools that wrap `rows.py`. Existing tools and `sync.py` are not modified.

## File Map

| File | Action | Responsibility |
|------|--------|----------------|
| `gssync/rows.py` | Create | Row filtering logic for GS and local files |
| `gssync/mcp_server.py` | Modify | Add 4 new MCP tools |
| `tests/test_rows.py` | Create | Unit tests for rows.py |
| `tests/test_mcp_server.py` | Modify | Tests for 4 new MCP tools |

## Row Filters

Exactly one filter must be provided per call. Passing more than one is an error.

| Parameter | Type | Example | Meaning |
|-----------|------|---------|---------|
| `row_numbers` | `str` | `"5"`, `"2,5,10"`, `"2:10"` | Row indices, 1-based (row 1 = header) |
| `filter_column` + `filter_value` | `str` + `str` | `"status"`, `"active"` | All rows where column value matches |
| `cell_range` | `str` | `"A2:D10"` | Google Sheets A1 notation |

`row_numbers` format rules:
- Single: `"5"` → row 5
- List: `"2,5,10"` → rows 2, 5, 10
- Range: `"2:10"` → rows 2 through 10 inclusive

Row indices are 1-based. Row 1 is always the header and is never included in results (but always included as the first row of output for context).

## New Tools

### `get_rows` — read rows from GS directly

Returns matching rows as a human-readable text table (header + data rows). No local file involved.

**Parameters:**
- `spreadsheet_url` *(required)*
- `sheet_name` *(required)*
- `row_numbers` | `filter_column`+`filter_value` | `cell_range` — exactly one filter

**Returns:** Formatted text, e.g.:
```
Sheet: Budget | Rows 3-5
name       | status | amount
-----------+--------+-------
Иван       | active | 1000
Мария      | done   | 2000
```

---

### `update_rows` — write rows to GS directly

Updates specific rows in GS without touching a local file.

**Parameters:**
- `spreadsheet_url` *(required)*
- `sheet_name` *(required)*
- `rows_data` *(required)* — JSON string, list of dicts keyed by header name:
  ```json
  [{"name": "Иван", "status": "done"}]
  ```
- `row_numbers` | `filter_column`+`filter_value` | `cell_range` — exactly one filter

**Write semantics:**
- `row_numbers` — replaces the specified rows with `rows_data` in order
- `filter_column`+`filter_value` — finds all matching rows in GS, replaces each with the first entry in `rows_data` (broadcast update: same values applied to all matches)
- `cell_range` — writes `rows_data` values into that exact range

**Returns:** `"Updated N rows in 'SheetName'"`

---

### `pull_rows` — download filtered rows into local file

Reads filtered rows from GS (including the header row) and writes them as the full sheet content in the local file, replacing any existing data for that sheet.

**Parameters:**
- `spreadsheet_url` *(required)*
- `sheet_name` *(required)*
- `file_path` *(required)*
- `file_format` — default `"xlsx"`
- `row_numbers` | `filter_column`+`filter_value` | `cell_range` — exactly one filter

**Returns:** `"Pulled N rows from 'SheetName' → file_path"`

---

### `push_rows` — upload filtered rows from local file to GS

Reads filtered rows from local file and updates corresponding positions in GS using the same write semantics as `update_rows`.

**Parameters:**
- `spreadsheet_url` *(required)*
- `sheet_name` *(required)*
- `file_path` *(required)*
- `file_format` — default `"xlsx"`
- `row_numbers` | `filter_column`+`filter_value` | `cell_range` — exactly one filter

**Returns:** `"Pushed N rows from file_path → 'SheetName'"`

---

## rows.py — Internal API

```python
# Row filter helpers
def parse_row_numbers(spec: str) -> list[int]: ...
# "5" → [5], "2,5,10" → [2,5,10], "2:10" → [2,3,...,10]

def filter_rows_by_column(rows: list[list], header: list, column: str, value: str) -> list[int]:
# Returns 1-based row indices (excluding header) that match

def parse_cell_range(cell_range: str) -> tuple[int, int, int, int]:
# "A2:D10" → (row_start, row_end, col_start, col_end), 1-based

# GS operations
def read_rows_from_sheet(spreadsheet, sheet_name: str, indices: list[int]) -> tuple[list, list[list]]:
# Returns (header, rows). indices are 1-based data rows (not counting header).

def write_rows_to_sheet(spreadsheet, sheet_name: str, indices: list[int], rows: list[list]) -> int:
# Writes rows to specific 1-based positions. Returns count of updated rows.

# Local file operations  
def read_rows_from_local(path: Path, fmt: str, sheet_name: str, indices: list[int]) -> tuple[list, list[list]]:
# Returns (header, rows)
```

## Error Handling

- More than one filter provided → `ValueError: "Provide exactly one filter: row_numbers, filter_column/filter_value, or cell_range"`
- No filter provided → same `ValueError`
- `row_numbers` out of bounds → silently skips missing rows, returns what exists
- `filter_column` not found in header → `ValueError: "Column 'X' not found in sheet"`
- Invalid `cell_range` format → `ValueError: "Invalid cell range: 'X'"`
- `rows_data` invalid JSON → `ValueError: "rows_data must be a JSON array of objects"`
- Sheet not found → existing `gspread.WorksheetNotFound` propagates as-is
