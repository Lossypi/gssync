# Formatting Sync Design

## Goal

Transfer cell and sheet formatting (font, color, fill, alignment, number format, borders, column widths, row heights) between Google Sheets and local xlsx files, in both directions, alongside the existing value sync.

## Scope

- **Directions:** both pull (GS→file) and push (file→GS)
- **Attributes:** full practical set — font (family/size/bold/italic/underline/strikethrough), text color, background fill, horizontal/vertical alignment + text wrap, column widths, row heights, number formats, borders. **No merged cells.**
- **File formats:** formatting is xlsx-only. For csv/json, formatting is silently skipped (values sync as today).
- **Tools:** whole-sheet tools (`pull_sheet`/`push_sheet`/`pull_all`/`push_all`) and row-level tools (`pull_rows`/`push_rows`). Row-level transfers per-cell formats + row heights, but **not** column widths (a sheet-level property).
- **Opt-out:** an `include_formatting: bool = True` parameter on every affected tool.

## Architecture (Approach A — parallel formatting layer)

The existing value pipeline (`Dict[str, List[List]]`) is left completely untouched — no regression risk to the 65 passing value-sync tests. Formatting is a separate, parallel concern read/written when `include_formatting` is true.

A new module `gssync/formatting.py` defines a backend-neutral intermediate representation (IR) plus pure conversion functions: Google Sheets JSON ↔ IR ↔ openpyxl styles. Nothing in `formatting.py` performs network or file I/O — all conversions are pure functions.

No new OAuth scope is required: the existing `https://www.googleapis.com/auth/spreadsheets` scope already covers reading and writing formatting.

## Data Model

```python
@dataclass
class CellFormat:
    font_family: str | None = None       # "Arial"
    font_size: float | None = None       # points, e.g. 11
    bold: bool | None = None
    italic: bool | None = None
    underline: bool | None = None
    strikethrough: bool | None = None
    text_color: str | None = None        # "#RRGGBB"
    background_color: str | None = None  # "#RRGGBB"
    h_align: str | None = None           # "left" | "center" | "right"
    v_align: str | None = None           # "top" | "middle" | "bottom"
    wrap: bool | None = None             # text wrap on/off
    number_format: str | None = None     # pattern, e.g. "0.00", "dd.mm.yyyy"
    borders: dict | None = None          # {"top": {"style","color"}, "bottom": {...}, ...}

@dataclass
class SheetFormatting:
    cells: dict[tuple[int, int], CellFormat]  # (row, col) 0-based → format
    column_widths: dict[int, int]             # col index (0-based) → pixels
    row_heights: dict[int, int]               # row index (0-based) → pixels
```

**Design rationale:**
- `None` for every attribute means "not set / inherit default". Only styled attributes are transferred; partial mappings degrade gracefully.
- Colors are normalized to `#RRGGBB` hex. GS uses 0–1 RGB floats, xlsx uses `FFRRGGBB` ARGB — the IR is the neutral middle.
- Widths and heights are stored in **pixels** (GS-native). The xlsx converter performs the lossy pixel↔Excel-unit conversion, so that logic lives in exactly one place.
- `cells` is sparse (only formatted cells present), keyed by 0-based `(row, col)` — the same structure works for full sheets and row subsets.

## Read/Write Mechanics

### Google Sheets side (`gssync/sheets.py`)

Uses gspread's existing API. No new scope.

**`read_sheet_formatting(spreadsheet, name) -> SheetFormatting`**
- Calls `spreadsheet.fetch_sheet_metadata(params={...})` with `includeGridData=True`, `ranges=[name]`, and a `fields` filter restricted to `sheets.data.rowData.values.userEnteredFormat`, `sheets.data.columnMetadata.pixelSize`, `sheets.data.rowMetadata.pixelSize`.
- Walks the returned grid; builds the IR via `formatting.gs_cell_to_ir()` per cell and reads column/row metadata for widths/heights.

**`write_sheet_formatting(spreadsheet, name, fmt: SheetFormatting) -> None`**
- `formatting.ir_to_gs_requests(sheet_id, fmt)` builds a list of `repeatCell` requests (per-cell formats) and `updateDimensionProperties` requests (column widths, row heights).
- Sent in a single `spreadsheet.batch_update({"requests": [...]})` call — one API round-trip regardless of cell count.
- The numeric `sheet_id` (not the title) is required for the requests; resolve it from worksheet metadata.

### xlsx side (`gssync/storage.py`)

Uses openpyxl, which stores styles natively.

**`read_xlsx_formatting(path, sheet_name) -> SheetFormatting`**
- Reads `cell.font`, `cell.fill`, `cell.alignment`, `cell.number_format`, `cell.border` and `ws.column_dimensions` / `ws.row_dimensions`.
- Builds the IR via `formatting.openpyxl_cell_to_ir()`.

**`apply_formatting_to_xlsx(path, sheet_name, fmt: SheetFormatting) -> None`**
- Opens the workbook, applies the IR via `formatting.ir_to_openpyxl_*()` helpers to cells and dimensions, saves.
- Assumes the sheet already exists (values written first by the existing `write_xlsx`).

### Conversion edge cases (all in `formatting.py`)

- **Color:** GS `{red, green, blue}` floats (0–1, missing channel = 0) ↔ `#RRGGBB` ↔ openpyxl `FFRRGGBB` ARGB.
- **Column width:** GS pixels ↔ Excel character units. `chars ≈ (pixels − 5) / 7`; reverse `pixels ≈ round(chars * 7 + 5)`.
- **Row height:** GS pixels ↔ openpyxl points. `points = pixels * 0.75`; reverse `pixels = round(points / 0.75)`.
- **Border styles:** mapping table between GS styles (`SOLID`, `SOLID_MEDIUM`, `SOLID_THICK`, `DASHED`, `DOTTED`, `DOUBLE`) and openpyxl styles (`thin`, `medium`, `thick`, `dashed`, `dotted`, `double`). Unknown styles are skipped.
- **Number format:** GS `numberFormat.pattern` ↔ openpyxl format code string. Both use ECMA-376 format codes, so patterns map directly; an empty/`NUMBER`-type "General" maps to openpyxl `"General"`.
- Any attribute that fails to map is left `None` / skipped, never raised.

## Tool/API Changes

### Sync layer (`gssync/sync.py`)

Each function gains `include_formatting: bool = True`:
- `pull_sheet` / `pull_all`: write values (unchanged); then if `include_formatting and fmt == "xlsx"`, read formatting from GS and apply to the xlsx file.
- `push_sheet` / `push_all`: push values (unchanged); then if `include_formatting and fmt == "xlsx"`, read formatting from xlsx and push to GS.

### Row layer (`gssync/rows.py`)

`pull_rows` / `push_rows` gain `include_formatting: bool = True`:
- Transfers per-cell formats for the filtered rows only (font/color/fill/align/number/border) plus row heights for those rows.
- **Does not** transfer column widths.
- Cell coordinates map through the existing 1-based-data-row → sheet-row offset (data row 1 = sheet row 2).

### MCP tools (`gssync/mcp_server.py`)

`pull_sheet`, `pull_all`, `push_sheet`, `push_all`, `pull_rows`, `push_rows` each get an `include_formatting: bool = True` parameter passed straight through to the underlying function. `get_rows` and `update_rows` are unchanged (value/text-oriented).

## Error Handling

- `include_formatting=True` with csv or json: skip formatting silently, append a note to the tool's return string, e.g. `"(formatting skipped: only xlsx supports it)"`.
- Per-attribute mapping failure: skip that attribute, continue with the rest.
- A formatting read/write failure does **not** roll back already-synced values. Values are committed independently; the formatting error is surfaced in the return message rather than raised, so a partial success is reported clearly.
- Unknown border styles, unparseable colors, or missing metadata fields are skipped, not raised.

## Testing

### `tests/test_formatting.py` (new) — the bulk of coverage

Pure-function unit tests, no network or real files:
- Color round-trips: GS floats → hex → ARGB → back; missing channels; clamping.
- Width conversion: pixels ↔ Excel char units, both directions, rounding.
- Row height conversion: pixels ↔ points, both directions.
- Border style mapping: each known style both directions; unknown style skipped.
- Number format pass-through and "General" handling.
- `gs_cell_to_ir` → `ir_to_gs_requests`: a styled cell produces the expected `repeatCell` request.
- `openpyxl_cell_to_ir` → `ir_to_openpyxl`: round-trip through an in-memory openpyxl workbook preserves attributes.
- `CellFormat` with all-`None` produces no GS request fields / no openpyxl mutation.

### Extend existing test files

- `tests/test_sync.py`: `include_formatting` wiring — formatting functions called for xlsx, skipped for csv/json (MagicMock spreadsheet, tmp_path file).
- `tests/test_rows.py`: `pull_rows`/`push_rows` with `include_formatting` — per-cell format transfer fires, column widths not touched.
- `tests/test_mcp_server.py`: the 6 tools accept and pass through `include_formatting`.

## File Map

| File | Action | Responsibility |
|------|--------|----------------|
| `gssync/formatting.py` | Create | IR dataclasses + all pure conversions (GS↔IR↔openpyxl, color/unit/border/number helpers) |
| `gssync/sheets.py` | Modify | `read_sheet_formatting` / `write_sheet_formatting` |
| `gssync/storage.py` | Modify | `read_xlsx_formatting` / `apply_formatting_to_xlsx` |
| `gssync/sync.py` | Modify | `include_formatting` param on `pull_sheet`/`pull_all`/`push_sheet`/`push_all` |
| `gssync/rows.py` | Modify | `include_formatting` on `pull_rows`/`push_rows` |
| `gssync/mcp_server.py` | Modify | thread `include_formatting` through 6 tools |
| `tests/test_formatting.py` | Create | pure-function unit tests |
| `tests/test_sync.py` | Modify | flag-wiring tests |
| `tests/test_rows.py` | Modify | row-level formatting tests |
| `tests/test_mcp_server.py` | Modify | tool param pass-through tests |
| `docs/GSSYNC_MCP.md` | Modify | document `include_formatting` |
| `README.md` | Modify | mention formatting sync |
