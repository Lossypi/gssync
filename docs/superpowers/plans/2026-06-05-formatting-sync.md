# Formatting Sync Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Transfer cell/sheet formatting (font, color, fill, alignment, number format, borders, column widths, row heights) between Google Sheets and local xlsx files in both directions, gated by an `include_formatting` flag.

**Architecture:** A new `gssync/formatting.py` holds a backend-neutral intermediate representation (IR) plus pure conversion functions (Google Sheets JSON ↔ IR ↔ openpyxl styles). The existing value pipeline is untouched; formatting is read/written as a parallel concern. `sheets.py`/`storage.py` get thin I/O wrappers; `sync.py`/`mcp_server.py` get an `include_formatting` flag; `rows.py` gets a pure row-remap helper.

**Tech Stack:** Python 3.10+, gspread (`fetch_sheet_metadata`, `batch_update`), openpyxl (`Font`/`PatternFill`/`Alignment`/`Border`/`Side`), pytest.

**Conventions for the implementing engineer:**
- Run Python with `py` (Windows). Run a single test file with `py -m pytest tests/test_X.py -v`.
- Coordinates in the IR are **0-based** `(row, col)`.
- Colors in the IR are `#RRGGBB`. GS uses 0–1 RGB floats; openpyxl uses 8-char `FFRRGGBB` ARGB.
- No new OAuth scope is needed — the existing `spreadsheets` scope covers formatting.

---

## File Map

| File | Action | Responsibility |
|------|--------|----------------|
| `gssync/formatting.py` | Create | IR dataclasses + all pure conversions (color/unit/border/number helpers, GS↔IR, openpyxl↔IR, request builder) |
| `gssync/sheets.py` | Modify | `read_sheet_formatting` / `write_sheet_formatting` (gspread I/O) |
| `gssync/storage.py` | Modify | `read_xlsx_formatting` / `apply_formatting_to_xlsx` (openpyxl I/O) |
| `gssync/sync.py` | Modify | `include_formatting` on `pull_sheet`/`pull_all`/`push_sheet`/`push_all` |
| `gssync/rows.py` | Modify | `remap_formatting_rows` pure helper |
| `gssync/mcp_server.py` | Modify | thread `include_formatting` through `pull_sheet`/`pull_all`/`push_sheet`/`push_all`/`pull_rows`/`push_rows` |
| `tests/test_formatting.py` | Create | pure-function unit tests |
| `tests/test_sheets_formatting.py` | Create | gspread I/O wrapper tests (MagicMock) |
| `tests/test_storage.py` | Create | xlsx formatting round-trip tests (tmp_path) |
| `tests/test_sync.py` | Modify | `include_formatting` wiring tests |
| `tests/test_rows.py` | Modify | `remap_formatting_rows` tests |
| `tests/test_mcp_server.py` | Modify | tool param pass-through tests |
| `docs/GSSYNC_MCP.md` | Modify | document `include_formatting` |
| `README.md` | Modify | mention formatting sync |

---

### Task 1: formatting.py — IR dataclasses and pure helpers

**Files:**
- Create: `gssync/formatting.py`
- Create: `tests/test_formatting.py`

- [ ] **Step 1: Write failing tests for dataclasses and helpers**

Create `tests/test_formatting.py`:

```python
import pytest

from gssync.formatting import (
    CellFormat,
    SheetFormatting,
    color_gs_to_hex,
    color_hex_to_gs,
    color_hex_to_argb,
    color_argb_to_hex,
    px_to_char_width,
    char_width_to_px,
    px_to_points,
    points_to_px,
    gs_number_type,
    GS_TO_OPENPYXL_BORDER,
    OPENPYXL_TO_GS_BORDER,
)


def test_cellformat_is_empty_true_for_all_none():
    assert CellFormat().is_empty() is True


def test_cellformat_is_empty_false_when_any_set():
    assert CellFormat(bold=True).is_empty() is False


def test_sheetformatting_defaults_are_independent():
    a = SheetFormatting()
    b = SheetFormatting()
    a.cells[(0, 0)] = CellFormat(bold=True)
    assert b.cells == {}


def test_color_gs_to_hex_full_channels():
    assert color_gs_to_hex({"red": 1.0, "green": 0.0, "blue": 0.0}) == "#FF0000"


def test_color_gs_to_hex_missing_channel_defaults_zero():
    assert color_gs_to_hex({"red": 1.0}) == "#FF0000"


def test_color_hex_to_gs_roundtrip():
    gs = color_hex_to_gs("#FF8000")
    assert gs["red"] == pytest.approx(1.0)
    assert gs["green"] == pytest.approx(128 / 255)
    assert gs["blue"] == pytest.approx(0.0)


def test_color_hex_to_argb():
    assert color_hex_to_argb("#1A2B3C") == "FF1A2B3C"


def test_color_argb_to_hex():
    assert color_argb_to_hex("FF1A2B3C") == "#1A2B3C"


def test_width_conversion_roundtrip_is_close():
    px = char_width_to_px(8.43)
    assert abs(px_to_char_width(px) - 8.43) < 0.5


def test_char_width_to_px_known_value():
    assert char_width_to_px(8.43) == round(8.43 * 7 + 5)


def test_height_conversion_roundtrip():
    assert points_to_px(px_to_points(20)) == 20


def test_gs_number_type_date():
    assert gs_number_type("dd.mm.yyyy") == "DATE"


def test_gs_number_type_percent():
    assert gs_number_type("0.0%") == "PERCENT"


def test_gs_number_type_plain_number():
    assert gs_number_type("0.00") == "NUMBER"


def test_border_maps_are_inverse_for_common_styles():
    assert GS_TO_OPENPYXL_BORDER["SOLID"] == "thin"
    assert OPENPYXL_TO_GS_BORDER["thin"] == "SOLID"
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `py -m pytest tests/test_formatting.py -v`
Expected: `ImportError` — `gssync.formatting` does not exist

- [ ] **Step 3: Create gssync/formatting.py with dataclasses and helpers**

Create `gssync/formatting.py`:

```python
"""Backend-neutral formatting intermediate representation (IR) and pure conversions.

No network or file I/O lives here — every function is pure and unit-testable.
Coordinates are 0-based (row, col). Colors are '#RRGGBB'.
"""
from dataclasses import dataclass, field
from typing import Optional


# ── IR dataclasses ──────────────────────────────────────────────────────────

@dataclass
class CellFormat:
    font_family: Optional[str] = None
    font_size: Optional[float] = None
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    underline: Optional[bool] = None
    strikethrough: Optional[bool] = None
    text_color: Optional[str] = None        # "#RRGGBB"
    background_color: Optional[str] = None  # "#RRGGBB"
    h_align: Optional[str] = None           # "left" | "center" | "right"
    v_align: Optional[str] = None           # "top" | "middle" | "bottom"
    wrap: Optional[bool] = None
    number_format: Optional[str] = None     # pattern string
    borders: Optional[dict] = None          # {"top": {"style","color"}, ...}

    def is_empty(self) -> bool:
        return all(getattr(self, name) is None for name in self.__dataclass_fields__)


@dataclass
class SheetFormatting:
    cells: dict = field(default_factory=dict)          # (row, col) -> CellFormat
    column_widths: dict = field(default_factory=dict)  # col -> pixels
    row_heights: dict = field(default_factory=dict)    # row -> pixels


# ── Color conversion ────────────────────────────────────────────────────────

def color_gs_to_hex(color: dict) -> str:
    r = round(color.get("red", 0.0) * 255)
    g = round(color.get("green", 0.0) * 255)
    b = round(color.get("blue", 0.0) * 255)
    return f"#{r:02X}{g:02X}{b:02X}"


def color_hex_to_gs(hex_str: str) -> dict:
    h = hex_str.lstrip("#")
    return {
        "red": int(h[0:2], 16) / 255,
        "green": int(h[2:4], 16) / 255,
        "blue": int(h[4:6], 16) / 255,
    }


def color_hex_to_argb(hex_str: str) -> str:
    return "FF" + hex_str.lstrip("#").upper()


def color_argb_to_hex(argb: str) -> str:
    return "#" + argb[-6:].upper()


# ── Dimension conversion ──────────────────────────────────────────────────────

def px_to_char_width(px: int) -> float:
    return round((px - 5) / 7, 2)


def char_width_to_px(chars: float) -> int:
    return round(chars * 7 + 5)


def px_to_points(px: int) -> float:
    return round(px * 0.75, 2)


def points_to_px(points: float) -> int:
    return round(points / 0.75)


# ── Number format type heuristic ──────────────────────────────────────────────

def gs_number_type(pattern: str) -> str:
    """Pick a Google Sheets numberFormat.type for an ECMA pattern string."""
    p = pattern.lower()
    has_date = ("yy" in p) or ("mmm" in p) or ("dd" in p) or ("yyyy" in p)
    has_time = ("h" in p and ":" in pattern) or ("ss" in p)
    if has_date and has_time:
        return "DATE_TIME"
    if has_date:
        return "DATE"
    if has_time:
        return "TIME"
    if "%" in pattern:
        return "PERCENT"
    if any(sym in pattern for sym in ("$", "€", "₽", "£")):
        return "CURRENCY"
    return "NUMBER"


# ── Border style maps ──────────────────────────────────────────────────────────

GS_TO_OPENPYXL_BORDER = {
    "SOLID": "thin",
    "SOLID_MEDIUM": "medium",
    "SOLID_THICK": "thick",
    "DASHED": "dashed",
    "DOTTED": "dotted",
    "DOUBLE": "double",
}

OPENPYXL_TO_GS_BORDER = {
    "thin": "SOLID",
    "medium": "SOLID_MEDIUM",
    "thick": "SOLID_THICK",
    "dashed": "DASHED",
    "dotted": "DOTTED",
    "double": "DOUBLE",
    "hair": "SOLID",
}

_GS_V_TO_IR = {"TOP": "top", "MIDDLE": "middle", "BOTTOM": "bottom"}
_IR_V_TO_GS = {"top": "TOP", "middle": "MIDDLE", "bottom": "BOTTOM"}
_OPENPYXL_V_TO_IR = {"top": "top", "center": "middle", "bottom": "bottom"}
_IR_V_TO_OPENPYXL = {"top": "top", "middle": "center", "bottom": "bottom"}
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `py -m pytest tests/test_formatting.py -v`
Expected: all PASS

- [ ] **Step 5: Commit**

```bash
git add gssync/formatting.py tests/test_formatting.py
git commit -m "feat: add formatting IR dataclasses and pure conversion helpers"
```

---

### Task 2: formatting.py — Google Sheets JSON ↔ IR

**Files:**
- Modify: `gssync/formatting.py`
- Modify: `tests/test_formatting.py`

- [ ] **Step 1: Append failing tests**

Append to `tests/test_formatting.py`:

```python
from gssync.formatting import (
    gs_cell_to_ir,
    ir_to_gs_cell_format,
    gs_grid_to_sheet_formatting,
    ir_to_gs_requests,
)


def test_gs_cell_to_ir_reads_text_format():
    value = {"userEnteredFormat": {"textFormat": {"bold": True, "fontSize": 14,
             "foregroundColor": {"red": 1.0}}}}
    cf = gs_cell_to_ir(value)
    assert cf.bold is True
    assert cf.font_size == 14
    assert cf.text_color == "#FF0000"


def test_gs_cell_to_ir_reads_align_wrap_fill():
    value = {"userEnteredFormat": {
        "backgroundColor": {"blue": 1.0},
        "horizontalAlignment": "CENTER",
        "verticalAlignment": "MIDDLE",
        "wrapStrategy": "WRAP",
    }}
    cf = gs_cell_to_ir(value)
    assert cf.background_color == "#0000FF"
    assert cf.h_align == "center"
    assert cf.v_align == "middle"
    assert cf.wrap is True


def test_gs_cell_to_ir_empty_value():
    assert gs_cell_to_ir({}).is_empty() is True


def test_ir_to_gs_cell_format_builds_fields():
    cf = CellFormat(bold=True, background_color="#0000FF")
    uef, fields = ir_to_gs_cell_format(cf)
    assert uef["textFormat"]["bold"] is True
    assert uef["backgroundColor"] == color_hex_to_gs("#0000FF")
    assert "textFormat.bold" in fields
    assert "backgroundColor" in fields


def test_ir_to_gs_cell_format_empty_has_no_fields():
    _, fields = ir_to_gs_cell_format(CellFormat())
    assert fields == []


def test_gs_grid_to_sheet_formatting_cells_and_dims():
    sheet = {"data": [{
        "rowData": [
            {"values": [{"userEnteredFormat": {"textFormat": {"bold": True}}}, {}]},
            {"values": [{}, {"userEnteredFormat": {"textFormat": {"italic": True}}}]},
        ],
        "columnMetadata": [{"pixelSize": 100}, {"pixelSize": 150}],
        "rowMetadata": [{"pixelSize": 21}, {"pixelSize": 30}],
    }]}
    sf = gs_grid_to_sheet_formatting(sheet)
    assert sf.cells[(0, 0)].bold is True
    assert sf.cells[(1, 1)].italic is True
    assert (0, 1) not in sf.cells
    assert sf.column_widths == {0: 100, 1: 150}
    assert sf.row_heights == {0: 21, 1: 30}


def test_ir_to_gs_requests_builds_repeatcell_and_dims():
    sf = SheetFormatting(
        cells={(2, 3): CellFormat(bold=True)},
        column_widths={1: 120},
        row_heights={4: 30},
    )
    requests = ir_to_gs_requests(99, sf)
    repeat = [r for r in requests if "repeatCell" in r][0]
    rng = repeat["repeatCell"]["range"]
    assert rng == {"sheetId": 99, "startRowIndex": 2, "endRowIndex": 3,
                   "startColumnIndex": 3, "endColumnIndex": 4}
    assert repeat["repeatCell"]["cell"]["userEnteredFormat"]["textFormat"]["bold"] is True
    dims = [r for r in requests if "updateDimensionProperties" in r]
    assert any(d["updateDimensionProperties"]["dimension"] == "COLUMNS" for d in dims)
    assert any(d["updateDimensionProperties"]["dimension"] == "ROWS" for d in dims)


def test_ir_to_gs_requests_skips_empty_cells():
    sf = SheetFormatting(cells={(0, 0): CellFormat()})
    assert ir_to_gs_requests(1, sf) == []
```

- [ ] **Step 2: Run to verify failure**

Run: `py -m pytest tests/test_formatting.py -v`
Expected: `ImportError` for `gs_cell_to_ir`

- [ ] **Step 3: Append GS↔IR conversions to gssync/formatting.py**

Append to `gssync/formatting.py`:

```python
# ── Google Sheets JSON ↔ IR ────────────────────────────────────────────────────

def gs_cell_to_ir(value: dict) -> CellFormat:
    fmt = value.get("userEnteredFormat", {})
    cf = CellFormat()
    tf = fmt.get("textFormat", {})
    cf.font_family = tf.get("fontFamily")
    cf.font_size = tf.get("fontSize")
    cf.bold = tf.get("bold")
    cf.italic = tf.get("italic")
    cf.underline = tf.get("underline")
    cf.strikethrough = tf.get("strikethrough")
    if "foregroundColor" in tf:
        cf.text_color = color_gs_to_hex(tf["foregroundColor"])
    if "backgroundColor" in fmt:
        cf.background_color = color_gs_to_hex(fmt["backgroundColor"])
    ha = fmt.get("horizontalAlignment")
    if ha:
        cf.h_align = ha.lower()
    va = fmt.get("verticalAlignment")
    if va:
        cf.v_align = _GS_V_TO_IR.get(va)
    wrap_strategy = fmt.get("wrapStrategy")
    if wrap_strategy is not None:
        cf.wrap = wrap_strategy == "WRAP"
    nf = fmt.get("numberFormat", {})
    if nf.get("pattern"):
        cf.number_format = nf["pattern"]
    borders = fmt.get("borders")
    if borders:
        mapped = {}
        for side in ("top", "bottom", "left", "right"):
            spec = borders.get(side)
            if not spec:
                continue
            style = GS_TO_OPENPYXL_BORDER.get(spec.get("style"))
            if not style:
                continue
            entry = {"style": style}
            if "color" in spec:
                entry["color"] = color_gs_to_hex(spec["color"])
            mapped[side] = entry
        if mapped:
            cf.borders = mapped
    return cf


def ir_to_gs_cell_format(cf: CellFormat):
    uef = {}
    fields = []
    text_format = {}
    if cf.font_family is not None:
        text_format["fontFamily"] = cf.font_family
        fields.append("textFormat.fontFamily")
    if cf.font_size is not None:
        text_format["fontSize"] = cf.font_size
        fields.append("textFormat.fontSize")
    if cf.bold is not None:
        text_format["bold"] = cf.bold
        fields.append("textFormat.bold")
    if cf.italic is not None:
        text_format["italic"] = cf.italic
        fields.append("textFormat.italic")
    if cf.underline is not None:
        text_format["underline"] = cf.underline
        fields.append("textFormat.underline")
    if cf.strikethrough is not None:
        text_format["strikethrough"] = cf.strikethrough
        fields.append("textFormat.strikethrough")
    if cf.text_color is not None:
        text_format["foregroundColor"] = color_hex_to_gs(cf.text_color)
        fields.append("textFormat.foregroundColor")
    if text_format:
        uef["textFormat"] = text_format
    if cf.background_color is not None:
        uef["backgroundColor"] = color_hex_to_gs(cf.background_color)
        fields.append("backgroundColor")
    if cf.h_align is not None:
        uef["horizontalAlignment"] = cf.h_align.upper()
        fields.append("horizontalAlignment")
    if cf.v_align is not None:
        uef["verticalAlignment"] = _IR_V_TO_GS.get(cf.v_align, cf.v_align.upper())
        fields.append("verticalAlignment")
    if cf.wrap is not None:
        uef["wrapStrategy"] = "WRAP" if cf.wrap else "OVERFLOW_CELL"
        fields.append("wrapStrategy")
    if cf.number_format is not None:
        uef["numberFormat"] = {"type": gs_number_type(cf.number_format),
                               "pattern": cf.number_format}
        fields.append("numberFormat")
    if cf.borders:
        gs_borders = {}
        for side, spec in cf.borders.items():
            gs_style = OPENPYXL_TO_GS_BORDER.get(spec.get("style"), "SOLID")
            entry = {"style": gs_style}
            if "color" in spec:
                entry["color"] = color_hex_to_gs(spec["color"])
            gs_borders[side] = entry
        uef["borders"] = gs_borders
        fields.append("borders")
    return uef, fields


def gs_grid_to_sheet_formatting(sheet: dict) -> SheetFormatting:
    sf = SheetFormatting()
    data_list = sheet.get("data", [])
    if not data_list:
        return sf
    data = data_list[0]
    for r, row in enumerate(data.get("rowData", [])):
        for c, value in enumerate(row.get("values", [])):
            cf = gs_cell_to_ir(value)
            if not cf.is_empty():
                sf.cells[(r, c)] = cf
    for c, meta in enumerate(data.get("columnMetadata", [])):
        if meta.get("pixelSize") is not None:
            sf.column_widths[c] = meta["pixelSize"]
    for r, meta in enumerate(data.get("rowMetadata", [])):
        if meta.get("pixelSize") is not None:
            sf.row_heights[r] = meta["pixelSize"]
    return sf


def ir_to_gs_requests(sheet_id: int, sf: SheetFormatting) -> list:
    requests = []
    for (row, col), cf in sf.cells.items():
        uef, fields = ir_to_gs_cell_format(cf)
        if not fields:
            continue
        requests.append({
            "repeatCell": {
                "range": {"sheetId": sheet_id, "startRowIndex": row, "endRowIndex": row + 1,
                          "startColumnIndex": col, "endColumnIndex": col + 1},
                "cell": {"userEnteredFormat": uef},
                "fields": "userEnteredFormat(" + ",".join(fields) + ")",
            }
        })
    for col, px in sf.column_widths.items():
        requests.append({
            "updateDimensionProperties": {
                "range": {"sheetId": sheet_id, "dimension": "COLUMNS",
                          "startIndex": col, "endIndex": col + 1},
                "properties": {"pixelSize": px},
                "fields": "pixelSize",
            }
        })
    for row, px in sf.row_heights.items():
        requests.append({
            "updateDimensionProperties": {
                "range": {"sheetId": sheet_id, "dimension": "ROWS",
                          "startIndex": row, "endIndex": row + 1},
                "properties": {"pixelSize": px},
                "fields": "pixelSize",
            }
        })
    return requests
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `py -m pytest tests/test_formatting.py -v`
Expected: all PASS

- [ ] **Step 5: Commit**

```bash
git add gssync/formatting.py tests/test_formatting.py
git commit -m "feat: add Google Sheets JSON <-> formatting IR conversions"
```

---

### Task 3: formatting.py — openpyxl ↔ IR

**Files:**
- Modify: `gssync/formatting.py`
- Modify: `tests/test_formatting.py`

- [ ] **Step 1: Append failing tests**

Append to `tests/test_formatting.py`:

```python
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

from gssync.formatting import openpyxl_cell_to_ir, apply_ir_to_openpyxl_cell


def test_openpyxl_cell_to_ir_reads_font_and_color():
    wb = openpyxl.Workbook()
    ws = wb.active
    cell = ws["A1"]
    cell.font = Font(name="Arial", size=14, bold=True, color="FFFF0000")
    cf = openpyxl_cell_to_ir(cell)
    assert cf.font_family == "Arial"
    assert cf.font_size == 14
    assert cf.bold is True
    assert cf.text_color == "#FF0000"


def test_openpyxl_cell_to_ir_reads_fill_align_wrap():
    wb = openpyxl.Workbook()
    ws = wb.active
    cell = ws["A1"]
    cell.fill = PatternFill(fill_type="solid", fgColor="FF00FF00")
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cf = openpyxl_cell_to_ir(cell)
    assert cf.background_color == "#00FF00"
    assert cf.h_align == "center"
    assert cf.v_align == "middle"
    assert cf.wrap is True


def test_openpyxl_cell_to_ir_skips_general_number_format():
    wb = openpyxl.Workbook()
    cf = openpyxl_cell_to_ir(wb.active["A1"])
    assert cf.number_format is None


def test_apply_ir_to_openpyxl_cell_roundtrip():
    wb = openpyxl.Workbook()
    ws = wb.active
    cell = ws["A1"]
    src = CellFormat(font_family="Calibri", font_size=12, bold=True, italic=True,
                     text_color="#112233", background_color="#445566",
                     h_align="right", v_align="top", wrap=True, number_format="0.00")
    apply_ir_to_openpyxl_cell(cell, src)
    back = openpyxl_cell_to_ir(cell)
    assert back.font_family == "Calibri"
    assert back.bold is True
    assert back.italic is True
    assert back.text_color == "#112233"
    assert back.background_color == "#445566"
    assert back.h_align == "right"
    assert back.v_align == "top"
    assert back.wrap is True
    assert back.number_format == "0.00"


def test_apply_ir_to_openpyxl_cell_empty_is_noop():
    wb = openpyxl.Workbook()
    cell = wb.active["A1"]
    apply_ir_to_openpyxl_cell(cell, CellFormat())
    assert cell.font.name in (None, "Calibri")  # default font untouched


def test_apply_ir_border_roundtrip():
    wb = openpyxl.Workbook()
    cell = wb.active["A1"]
    apply_ir_to_openpyxl_cell(cell, CellFormat(borders={"top": {"style": "thin", "color": "#000000"}}))
    back = openpyxl_cell_to_ir(cell)
    assert back.borders["top"]["style"] == "thin"
```

- [ ] **Step 2: Run to verify failure**

Run: `py -m pytest tests/test_formatting.py -v`
Expected: `ImportError` for `openpyxl_cell_to_ir`

- [ ] **Step 3: Append openpyxl↔IR conversions to gssync/formatting.py**

First, add this import at the **top** of `gssync/formatting.py` (after the existing `from typing import Optional`):

```python
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
```

Then append to the end of `gssync/formatting.py`:

```python
# ── openpyxl ↔ IR ──────────────────────────────────────────────────────────────

def _openpyxl_color_rgb(color) -> Optional[str]:
    if color is None:
        return None
    rgb = getattr(color, "rgb", None)
    if isinstance(rgb, str) and len(rgb) == 8:
        return rgb
    return None


def openpyxl_cell_to_ir(cell) -> CellFormat:
    cf = CellFormat()
    font = cell.font
    if font is not None:
        if font.name is not None:
            cf.font_family = font.name
        if font.size is not None:
            cf.font_size = float(font.size)
        if font.bold is not None:
            cf.bold = bool(font.bold)
        if font.italic is not None:
            cf.italic = bool(font.italic)
        if font.underline:
            cf.underline = True
        if font.strike is not None:
            cf.strikethrough = bool(font.strike)
        rgb = _openpyxl_color_rgb(font.color)
        if rgb:
            cf.text_color = color_argb_to_hex(rgb)
    fill = cell.fill
    if fill is not None and getattr(fill, "fill_type", None) == "solid":
        rgb = _openpyxl_color_rgb(fill.fgColor)
        if rgb:
            cf.background_color = color_argb_to_hex(rgb)
    al = cell.alignment
    if al is not None:
        if al.horizontal in ("left", "center", "right"):
            cf.h_align = al.horizontal
        if al.vertical in _OPENPYXL_V_TO_IR:
            cf.v_align = _OPENPYXL_V_TO_IR[al.vertical]
        if al.wrap_text is not None:
            cf.wrap = bool(al.wrap_text)
    nf = cell.number_format
    if nf and nf != "General":
        cf.number_format = nf
    border = cell.border
    if border is not None:
        mapped = {}
        for side_name in ("top", "bottom", "left", "right"):
            side = getattr(border, side_name)
            if side is not None and side.style:
                entry = {"style": side.style}
                rgb = _openpyxl_color_rgb(side.color)
                if rgb:
                    entry["color"] = color_argb_to_hex(rgb)
                mapped[side_name] = entry
        if mapped:
            cf.borders = mapped
    return cf


def apply_ir_to_openpyxl_cell(cell, cf: CellFormat) -> None:
    if cf.is_empty():
        return
    font_kwargs = {}
    if cf.font_family is not None:
        font_kwargs["name"] = cf.font_family
    if cf.font_size is not None:
        font_kwargs["size"] = cf.font_size
    if cf.bold is not None:
        font_kwargs["bold"] = cf.bold
    if cf.italic is not None:
        font_kwargs["italic"] = cf.italic
    if cf.underline is not None:
        font_kwargs["underline"] = "single" if cf.underline else None
    if cf.strikethrough is not None:
        font_kwargs["strike"] = cf.strikethrough
    if cf.text_color is not None:
        font_kwargs["color"] = color_hex_to_argb(cf.text_color)
    if font_kwargs:
        cell.font = Font(**font_kwargs)
    if cf.background_color is not None:
        cell.fill = PatternFill(fill_type="solid", fgColor=color_hex_to_argb(cf.background_color))
    align_kwargs = {}
    if cf.h_align is not None:
        align_kwargs["horizontal"] = cf.h_align
    if cf.v_align is not None:
        align_kwargs["vertical"] = _IR_V_TO_OPENPYXL.get(cf.v_align, cf.v_align)
    if cf.wrap is not None:
        align_kwargs["wrap_text"] = cf.wrap
    if align_kwargs:
        cell.alignment = Alignment(**align_kwargs)
    if cf.number_format is not None:
        cell.number_format = cf.number_format
    if cf.borders:
        sides = {}
        for side_name, spec in cf.borders.items():
            side_kwargs = {"style": spec.get("style")}
            if "color" in spec:
                side_kwargs["color"] = color_hex_to_argb(spec["color"])
            sides[side_name] = Side(**side_kwargs)
        cell.border = Border(**sides)
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `py -m pytest tests/test_formatting.py -v`
Expected: all PASS

- [ ] **Step 5: Commit**

```bash
git add gssync/formatting.py tests/test_formatting.py
git commit -m "feat: add openpyxl <-> formatting IR conversions"
```

---

### Task 4: sheets.py — read/write sheet formatting (gspread I/O)

**Files:**
- Modify: `gssync/sheets.py`
- Create: `tests/test_sheets_formatting.py`

- [ ] **Step 1: Write failing tests**

Create `tests/test_sheets_formatting.py`:

```python
from unittest.mock import MagicMock

from gssync.formatting import SheetFormatting, CellFormat
from gssync.sheets import read_sheet_formatting, write_sheet_formatting


def test_read_sheet_formatting_parses_metadata():
    mock_ss = MagicMock()
    mock_ss.fetch_sheet_metadata.return_value = {
        "sheets": [{
            "properties": {"sheetId": 7, "title": "Sheet1"},
            "data": [{
                "rowData": [{"values": [{"userEnteredFormat": {"textFormat": {"bold": True}}}]}],
                "columnMetadata": [{"pixelSize": 120}],
                "rowMetadata": [{"pixelSize": 21}],
            }],
        }]
    }
    sf = read_sheet_formatting(mock_ss, "Sheet1")
    assert sf.cells[(0, 0)].bold is True
    assert sf.column_widths == {0: 120}
    mock_ss.fetch_sheet_metadata.assert_called_once()


def test_read_sheet_formatting_missing_sheet_returns_empty():
    mock_ss = MagicMock()
    mock_ss.fetch_sheet_metadata.return_value = {"sheets": []}
    sf = read_sheet_formatting(mock_ss, "Nope")
    assert sf.cells == {}


def test_write_sheet_formatting_calls_batch_update():
    mock_ws = MagicMock()
    mock_ws.id = 7
    mock_ss = MagicMock()
    mock_ss.worksheet.return_value = mock_ws
    sf = SheetFormatting(cells={(0, 0): CellFormat(bold=True)})
    write_sheet_formatting(mock_ss, "Sheet1", sf)
    mock_ss.batch_update.assert_called_once()
    body = mock_ss.batch_update.call_args[0][0]
    assert "requests" in body
    assert body["requests"][0]["repeatCell"]["range"]["sheetId"] == 7


def test_write_sheet_formatting_no_requests_skips_call():
    mock_ws = MagicMock()
    mock_ws.id = 7
    mock_ss = MagicMock()
    mock_ss.worksheet.return_value = mock_ws
    write_sheet_formatting(mock_ss, "Sheet1", SheetFormatting())
    mock_ss.batch_update.assert_not_called()
```

- [ ] **Step 2: Run to verify failure**

Run: `py -m pytest tests/test_sheets_formatting.py -v`
Expected: `ImportError` for `read_sheet_formatting`

- [ ] **Step 3: Add functions to gssync/sheets.py**

Append to `gssync/sheets.py`:

```python
def read_sheet_formatting(spreadsheet: gspread.Spreadsheet, name: str):
    from .formatting import gs_grid_to_sheet_formatting, SheetFormatting
    fields = (
        "sheets(properties(sheetId,title),"
        "data(rowData(values(userEnteredFormat)),"
        "columnMetadata(pixelSize),rowMetadata(pixelSize)))"
    )
    meta = spreadsheet.fetch_sheet_metadata(params={
        "includeGridData": True,
        "ranges": [name],
        "fields": fields,
    })
    for sheet in meta.get("sheets", []):
        if sheet.get("properties", {}).get("title") == name:
            return gs_grid_to_sheet_formatting(sheet)
    return SheetFormatting()


def write_sheet_formatting(spreadsheet: gspread.Spreadsheet, name: str, sf) -> None:
    from .formatting import ir_to_gs_requests
    sheet_id = spreadsheet.worksheet(name).id
    requests = ir_to_gs_requests(sheet_id, sf)
    if requests:
        spreadsheet.batch_update({"requests": requests})
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `py -m pytest tests/test_sheets_formatting.py -v`
Expected: all PASS

- [ ] **Step 5: Commit**

```bash
git add gssync/sheets.py tests/test_sheets_formatting.py
git commit -m "feat: add read/write_sheet_formatting gspread wrappers"
```

---

### Task 5: storage.py — read/apply xlsx formatting (openpyxl I/O)

**Files:**
- Modify: `gssync/storage.py`
- Create: `tests/test_storage.py`

- [ ] **Step 1: Write failing tests**

Create `tests/test_storage.py`:

```python
import openpyxl
from openpyxl.styles import Font, PatternFill

from gssync.formatting import SheetFormatting, CellFormat
from gssync.storage import read_xlsx_formatting, apply_formatting_to_xlsx


def _make_xlsx(path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"] = "header"
    ws["A1"].font = Font(bold=True, color="FFFF0000")
    ws.column_dimensions["A"].width = 20
    ws.row_dimensions[1].height = 30
    wb.save(path)


def test_read_xlsx_formatting_reads_cells_and_dims(tmp_path):
    xlsx = tmp_path / "data.xlsx"
    _make_xlsx(xlsx)
    sf = read_xlsx_formatting(xlsx, "Sheet1")
    assert sf.cells[(0, 0)].bold is True
    assert sf.cells[(0, 0)].text_color == "#FF0000"
    assert 0 in sf.column_widths
    assert 0 in sf.row_heights


def test_read_xlsx_formatting_missing_sheet_returns_empty(tmp_path):
    xlsx = tmp_path / "data.xlsx"
    _make_xlsx(xlsx)
    sf = read_xlsx_formatting(xlsx, "Missing")
    assert sf.cells == {}


def test_apply_formatting_to_xlsx_writes_styles(tmp_path):
    xlsx = tmp_path / "data.xlsx"
    wb = openpyxl.Workbook()
    wb.active.title = "Sheet1"
    wb.active["A1"] = "x"
    wb.save(xlsx)
    sf = SheetFormatting(
        cells={(0, 0): CellFormat(bold=True, background_color="#00FF00")},
        column_widths={0: 150},
        row_heights={0: 40},
    )
    apply_formatting_to_xlsx(xlsx, "Sheet1", sf)
    back = read_xlsx_formatting(xlsx, "Sheet1")
    assert back.cells[(0, 0)].bold is True
    assert back.cells[(0, 0)].background_color == "#00FF00"
    assert 0 in back.column_widths
```

- [ ] **Step 2: Run to verify failure**

Run: `py -m pytest tests/test_storage.py -v`
Expected: `ImportError` for `read_xlsx_formatting`

- [ ] **Step 3: Add functions to gssync/storage.py**

Add this import near the top of `gssync/storage.py` (after `import openpyxl`):

```python
from openpyxl.utils import get_column_letter, column_index_from_string
```

Append to `gssync/storage.py`:

```python
def read_xlsx_formatting(path: Path, sheet_name: str):
    from .formatting import (
        openpyxl_cell_to_ir, char_width_to_px, points_to_px, SheetFormatting,
    )
    wb = openpyxl.load_workbook(path)
    if sheet_name not in wb.sheetnames:
        return SheetFormatting()
    ws = wb[sheet_name]
    sf = SheetFormatting()
    for row in ws.iter_rows():
        for cell in row:
            # Skip cells that were never explicitly styled. openpyxl hands every
            # cell a default Font(Calibri, 11), so without this guard every blank
            # cell would be captured and pushed to GS.
            if not cell.has_style:
                continue
            cf = openpyxl_cell_to_ir(cell)
            if not cf.is_empty():
                sf.cells[(cell.row - 1, cell.column - 1)] = cf
    for letter, dim in ws.column_dimensions.items():
        if dim.width is not None:
            sf.column_widths[column_index_from_string(letter) - 1] = char_width_to_px(dim.width)
    for idx, dim in ws.row_dimensions.items():
        if dim.height is not None:
            sf.row_heights[idx - 1] = points_to_px(dim.height)
    return sf


def apply_formatting_to_xlsx(path: Path, sheet_name: str, sf) -> None:
    from .formatting import apply_ir_to_openpyxl_cell, px_to_char_width, px_to_points
    wb = openpyxl.load_workbook(path)
    ws = wb[sheet_name]
    for (r, c), cf in sf.cells.items():
        apply_ir_to_openpyxl_cell(ws.cell(row=r + 1, column=c + 1), cf)
    for col, px in sf.column_widths.items():
        ws.column_dimensions[get_column_letter(col + 1)].width = px_to_char_width(px)
    for row, px in sf.row_heights.items():
        ws.row_dimensions[row + 1].height = px_to_points(px)
    wb.save(path)
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `py -m pytest tests/test_storage.py -v`
Expected: all PASS

- [ ] **Step 5: Commit**

```bash
git add gssync/storage.py tests/test_storage.py
git commit -m "feat: add read/apply xlsx formatting via openpyxl"
```

---

### Task 6: sync.py — include_formatting on whole-sheet sync

**Files:**
- Modify: `gssync/sync.py`
- Modify: `tests/test_sync.py`

- [ ] **Step 1: Add failing tests to tests/test_sync.py**

Append to `tests/test_sync.py` (imports `MagicMock`, `patch`, `Path` are already present in that file; if not, add `from unittest.mock import MagicMock, patch`):

```python
from unittest.mock import patch


def test_pull_sheet_applies_formatting_for_xlsx(tmp_path):
    from gssync.sync import pull_sheet
    mock_ss = MagicMock()
    xlsx = tmp_path / "data.xlsx"
    with patch("gssync.sync.read_sheet", return_value=[["a", "b"], ["1", "2"]]), \
         patch("gssync.sheets.read_sheet_formatting") as mock_read_fmt, \
         patch("gssync.storage.apply_formatting_to_xlsx") as mock_apply:
        pull_sheet(mock_ss, "Sheet1", xlsx, "xlsx", include_formatting=True)
    mock_read_fmt.assert_called_once_with(mock_ss, "Sheet1")
    mock_apply.assert_called_once()


def test_pull_sheet_skips_formatting_for_csv(tmp_path):
    from gssync.sync import pull_sheet
    mock_ss = MagicMock()
    csv_dir = tmp_path / "out"
    with patch("gssync.sync.read_sheet", return_value=[["a"]]), \
         patch("gssync.sheets.read_sheet_formatting") as mock_read_fmt:
        pull_sheet(mock_ss, "Sheet1", csv_dir, "csv", include_formatting=True)
    mock_read_fmt.assert_not_called()


def test_pull_sheet_skips_formatting_when_flag_off(tmp_path):
    from gssync.sync import pull_sheet
    mock_ss = MagicMock()
    xlsx = tmp_path / "data.xlsx"
    with patch("gssync.sync.read_sheet", return_value=[["a"]]), \
         patch("gssync.sheets.read_sheet_formatting") as mock_read_fmt:
        pull_sheet(mock_ss, "Sheet1", xlsx, "xlsx", include_formatting=False)
    mock_read_fmt.assert_not_called()


def test_push_sheet_writes_formatting_for_xlsx(tmp_path):
    from gssync.sync import push_sheet
    from gssync.storage import write_local
    mock_ss = MagicMock()
    xlsx = tmp_path / "data.xlsx"
    write_local(xlsx, "xlsx", {"Sheet1": [["a", "b"]]})
    with patch("gssync.sheets.write_sheet_formatting") as mock_write_fmt, \
         patch("gssync.storage.read_xlsx_formatting", return_value="SF"):
        push_sheet(mock_ss, "Sheet1", xlsx, "xlsx", include_formatting=True)
    mock_write_fmt.assert_called_once_with(mock_ss, "Sheet1", "SF")
```

- [ ] **Step 2: Run to verify failure**

Run: `py -m pytest tests/test_sync.py -v`
Expected: new tests FAIL (`pull_sheet` has no `include_formatting` param)

- [ ] **Step 3: Update gssync/sync.py**

Replace the entire body of `gssync/sync.py` with:

```python
from pathlib import Path

import gspread

from .sheets import list_sheet_names, read_sheet, write_sheet
from .storage import list_local_sheets, read_local, write_local


def _load_existing(path: Path, fmt: str) -> dict:
    exists = (path.exists() and path.is_dir()) if fmt == "csv" else path.exists()
    return read_local(path, fmt) if exists else {}


def _apply_pull_formatting(spreadsheet, name: str, file_path: Path, fmt: str) -> None:
    if fmt != "xlsx":
        return
    from .sheets import read_sheet_formatting
    from .storage import apply_formatting_to_xlsx
    sf = read_sheet_formatting(spreadsheet, name)
    apply_formatting_to_xlsx(file_path, name, sf)


def _push_formatting(spreadsheet, name: str, file_path: Path, fmt: str) -> None:
    if fmt != "xlsx":
        return
    from .sheets import write_sheet_formatting
    from .storage import read_xlsx_formatting
    sf = read_xlsx_formatting(file_path, name)
    write_sheet_formatting(spreadsheet, name, sf)


def pull_sheet(spreadsheet: gspread.Spreadsheet, sheet_name: str, file_path: Path,
               fmt: str, include_formatting: bool = True) -> None:
    data = read_sheet(spreadsheet, sheet_name)
    existing = _load_existing(file_path, fmt)
    existing[sheet_name] = data
    write_local(file_path, fmt, existing)
    if include_formatting:
        _apply_pull_formatting(spreadsheet, sheet_name, file_path, fmt)


def pull_all(spreadsheet: gspread.Spreadsheet, file_path: Path, fmt: str,
             include_formatting: bool = True) -> None:
    names = list_sheet_names(spreadsheet)
    existing = _load_existing(file_path, fmt)
    for name in names:
        existing[name] = read_sheet(spreadsheet, name)
    write_local(file_path, fmt, existing)
    if include_formatting:
        for name in names:
            _apply_pull_formatting(spreadsheet, name, file_path, fmt)


def push_sheet(spreadsheet: gspread.Spreadsheet, sheet_name: str, file_path: Path,
               fmt: str, include_formatting: bool = True) -> None:
    local_data = read_local(file_path, fmt)
    write_sheet(spreadsheet, sheet_name, local_data[sheet_name])
    if include_formatting:
        _push_formatting(spreadsheet, sheet_name, file_path, fmt)


def push_all(spreadsheet: gspread.Spreadsheet, file_path: Path, fmt: str,
             include_formatting: bool = True) -> None:
    local_data = read_local(file_path, fmt)
    for name, rows in local_data.items():
        write_sheet(spreadsheet, name, rows)
    if include_formatting:
        for name in local_data:
            _push_formatting(spreadsheet, name, file_path, fmt)
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `py -m pytest tests/test_sync.py -v`
Expected: all PASS (existing + new)

- [ ] **Step 5: Commit**

```bash
git add gssync/sync.py tests/test_sync.py
git commit -m "feat: add include_formatting to whole-sheet sync functions"
```

---

### Task 7: rows.py — remap_formatting_rows helper

**Files:**
- Modify: `gssync/rows.py`
- Modify: `tests/test_rows.py`

**Context:** `pull_rows`/`push_rows` MCP tools (Task 8) need to move per-cell formatting between original sheet rows and compacted file rows. This pure helper does the row remap. Row keys are 0-based. Column widths are intentionally dropped (sheet-level, meaningless for a row subset).

- [ ] **Step 1: Add failing tests to tests/test_rows.py**

Append to `tests/test_rows.py`:

```python
from gssync.formatting import SheetFormatting, CellFormat
from gssync.rows import remap_formatting_rows


def test_remap_formatting_rows_remaps_cells_and_heights():
    sf = SheetFormatting(
        cells={(0, 0): CellFormat(bold=True), (3, 1): CellFormat(italic=True)},
        column_widths={0: 100},
        row_heights={0: 21, 3: 40},
    )
    out = remap_formatting_rows(sf, {0: 0, 3: 1})
    assert out.cells[(0, 0)].bold is True
    assert out.cells[(1, 1)].italic is True
    assert out.row_heights == {0: 21, 1: 40}
    assert out.column_widths == {}  # dropped


def test_remap_formatting_rows_skips_unmapped():
    sf = SheetFormatting(cells={(5, 0): CellFormat(bold=True)})
    out = remap_formatting_rows(sf, {0: 0})
    assert out.cells == {}
```

- [ ] **Step 2: Run to verify failure**

Run: `py -m pytest tests/test_rows.py -v`
Expected: `ImportError` for `remap_formatting_rows`

- [ ] **Step 3: Add function to gssync/rows.py**

Append to `gssync/rows.py`:

```python
def remap_formatting_rows(sf, row_map: dict):
    """Remap a SheetFormatting onto new 0-based row positions.

    row_map maps source 0-based row -> destination 0-based row. Cells and row
    heights whose source row is in row_map are copied to the destination row;
    everything else is dropped. Column widths are always dropped (a sheet-level
    property that does not apply to a row subset).
    """
    from .formatting import SheetFormatting
    out = SheetFormatting()
    for (r, c), cf in sf.cells.items():
        if r in row_map:
            out.cells[(row_map[r], c)] = cf
    for r, px in sf.row_heights.items():
        if r in row_map:
            out.row_heights[row_map[r]] = px
    return out
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `py -m pytest tests/test_rows.py -v`
Expected: all PASS

- [ ] **Step 5: Commit**

```bash
git add gssync/rows.py tests/test_rows.py
git commit -m "feat: add remap_formatting_rows helper for row-level formatting"
```

---

### Task 8: mcp_server.py — thread include_formatting through tools + docs

**Files:**
- Modify: `gssync/mcp_server.py`
- Modify: `tests/test_mcp_server.py`
- Modify: `docs/GSSYNC_MCP.md`
- Modify: `README.md`

**Context — current MCP tool signatures (do not break value behavior):**
- `pull_sheet(spreadsheet_url, sheet_name, file_path, file_format="xlsx")` → calls `_pull_sheet(spreadsheet, sheet_name, Path(file_path), file_format)`
- `pull_all`, `push_sheet`, `push_all` mirror this.
- `pull_rows(spreadsheet_url, sheet_name, file_path, file_format="xlsx", row_numbers="", filter_column="", filter_value="", cell_range="")` — reads filtered rows from GS, writes `[header] + rows` to the file.
- `push_rows(...)` same params — reads filtered rows from the local file, writes them to GS at `sheet_row = idx + 1`.

**Row coordinate facts (used below):**
- A 1-based data row index `idx` corresponds to 0-based sheet row `idx` (header is 0-based row 0).
- In `pull_rows`, the output file is compacted: header → file 0-based row 0; the k-th selected row (k starting at 1) → file 0-based row k.
- In `push_rows`, file and GS share the same 0-based row for each `idx` (no compaction).

- [ ] **Step 1: Add failing tests to tests/test_mcp_server.py**

Append to `tests/test_mcp_server.py`:

```python
def test_pull_sheet_passes_include_formatting():
    mock_spreadsheet = MagicMock()
    with patch("gssync.mcp_server.get_client"), \
         patch("gssync.mcp_server.open_spreadsheet", return_value=mock_spreadsheet), \
         patch("gssync.mcp_server._pull_sheet") as mock_pull:
        pull_sheet(
            spreadsheet_url="https://docs.google.com/spreadsheets/d/abc",
            sheet_name="Sheet1",
            file_path="C:\\data\\report.xlsx",
            include_formatting=False,
        )
    assert mock_pull.call_args.kwargs["include_formatting"] is False


def test_push_sheet_passes_include_formatting():
    mock_spreadsheet = MagicMock()
    with patch("gssync.mcp_server.get_client"), \
         patch("gssync.mcp_server.open_spreadsheet", return_value=mock_spreadsheet), \
         patch("gssync.mcp_server._push_sheet") as mock_push:
        push_sheet(
            spreadsheet_url="https://docs.google.com/spreadsheets/d/abc",
            sheet_name="Sheet1",
            file_path="C:\\data\\report.xlsx",
        )
    assert mock_push.call_args.kwargs["include_formatting"] is True


def test_pull_rows_applies_formatting_for_xlsx(tmp_path):
    mock_ws = MagicMock()
    mock_ws.get_all_values.return_value = [
        ["name", "status"], ["Ivan", "active"], ["Maria", "done"],
    ]
    mock_ss = MagicMock()
    mock_ss.worksheet.return_value = mock_ws
    xlsx = tmp_path / "data.xlsx"
    from gssync.formatting import SheetFormatting, CellFormat
    sf = SheetFormatting(cells={(0, 0): CellFormat(bold=True), (1, 0): CellFormat(italic=True)})
    with patch("gssync.mcp_server.get_client"), \
         patch("gssync.mcp_server.open_spreadsheet", return_value=mock_ss), \
         patch("gssync.mcp_server.read_sheet_formatting", return_value=sf), \
         patch("gssync.mcp_server.apply_formatting_to_xlsx") as mock_apply:
        pull_rows(
            spreadsheet_url="https://docs.google.com/spreadsheets/d/abc",
            sheet_name="Sheet1",
            file_path=str(xlsx),
            row_numbers="1",
            include_formatting=True,
        )
    mock_apply.assert_called_once()
    remapped = mock_apply.call_args[0][2]
    # header row 0 -> file row 0; data row idx=1 (0-based sheet row 1) -> file row 1
    assert (0, 0) in remapped.cells
    assert (1, 0) in remapped.cells


def test_push_rows_writes_formatting_for_xlsx(tmp_path):
    from gssync.storage import write_local
    xlsx = tmp_path / "data.xlsx"
    write_local(xlsx, "xlsx", {"Sheet1": [["name", "status"], ["Ivan", "active"], ["Maria", "done"]]})
    mock_ws = MagicMock()
    mock_ss = MagicMock()
    mock_ss.worksheet.return_value = mock_ws
    from gssync.formatting import SheetFormatting, CellFormat
    sf = SheetFormatting(cells={(2, 0): CellFormat(bold=True)})
    with patch("gssync.mcp_server.get_client"), \
         patch("gssync.mcp_server.open_spreadsheet", return_value=mock_ss), \
         patch("gssync.mcp_server.read_xlsx_formatting", return_value=sf), \
         patch("gssync.mcp_server.write_sheet_formatting") as mock_write_fmt:
        push_rows(
            spreadsheet_url="https://docs.google.com/spreadsheets/d/abc",
            sheet_name="Sheet1",
            file_path=str(xlsx),
            row_numbers="2",
            include_formatting=True,
        )
    mock_write_fmt.assert_called_once()
    remapped = mock_write_fmt.call_args[0][2]
    # data row idx=2 -> 0-based row 2 on both sides (identity)
    assert (2, 0) in remapped.cells


def test_pull_rows_skips_formatting_for_csv(tmp_path):
    mock_ws = MagicMock()
    mock_ws.get_all_values.return_value = [["name"], ["Ivan"]]
    mock_ss = MagicMock()
    mock_ss.worksheet.return_value = mock_ws
    out_dir = tmp_path / "out"
    with patch("gssync.mcp_server.get_client"), \
         patch("gssync.mcp_server.open_spreadsheet", return_value=mock_ss), \
         patch("gssync.mcp_server.read_sheet_formatting") as mock_read_fmt:
        pull_rows(
            spreadsheet_url="https://docs.google.com/spreadsheets/d/abc",
            sheet_name="Sheet1",
            file_path=str(out_dir),
            file_format="csv",
            row_numbers="1",
            include_formatting=True,
        )
    mock_read_fmt.assert_not_called()
```

Also update the `from gssync.mcp_server import (...)` block at the top of the file is unchanged (all names already imported). No new tool names are added.

- [ ] **Step 2: Run to verify failure**

Run: `py -m pytest tests/test_mcp_server.py -v`
Expected: new tests FAIL (tools have no `include_formatting`; `read_sheet_formatting` etc. not imported in mcp_server)

- [ ] **Step 3: Add imports to gssync/mcp_server.py**

Add to the import block of `gssync/mcp_server.py` (after the existing `from .rows import (...)` block):

```python
from .rows import remap_formatting_rows
from .sheets import read_sheet_formatting, write_sheet_formatting
from .storage import apply_formatting_to_xlsx, read_xlsx_formatting
```

- [ ] **Step 4: Add include_formatting to the 4 whole-sheet tools**

In `gssync/mcp_server.py`, replace the four whole-sheet tools with these versions (add the param, pass it through as a keyword arg):

```python
@mcp.tool()
def pull_sheet(
    spreadsheet_url: str,
    sheet_name: str,
    file_path: str,
    file_format: str = "xlsx",
    include_formatting: bool = True,
) -> str:
    client = get_client()
    spreadsheet = open_spreadsheet(client, spreadsheet_url)
    _pull_sheet(spreadsheet, sheet_name, Path(file_path), file_format,
                include_formatting=include_formatting)
    return f"Pulled '{sheet_name}' from {spreadsheet_url} → {file_path}"


@mcp.tool()
def pull_all(
    spreadsheet_url: str,
    file_path: str,
    file_format: str = "xlsx",
    include_formatting: bool = True,
) -> str:
    client = get_client()
    spreadsheet = open_spreadsheet(client, spreadsheet_url)
    names = list_sheet_names(spreadsheet)
    _pull_all(spreadsheet, Path(file_path), file_format,
              include_formatting=include_formatting)
    return f"Pulled {len(names)} sheets from {spreadsheet_url} → {file_path}"


@mcp.tool()
def push_sheet(
    spreadsheet_url: str,
    sheet_name: str,
    file_path: str,
    file_format: str = "xlsx",
    include_formatting: bool = True,
) -> str:
    client = get_client()
    spreadsheet = open_spreadsheet(client, spreadsheet_url)
    _push_sheet(spreadsheet, sheet_name, Path(file_path), file_format,
                include_formatting=include_formatting)
    return f"Pushed '{sheet_name}' from {file_path} → {spreadsheet_url}"


@mcp.tool()
def push_all(
    spreadsheet_url: str,
    file_path: str,
    file_format: str = "xlsx",
    include_formatting: bool = True,
) -> str:
    client = get_client()
    spreadsheet = open_spreadsheet(client, spreadsheet_url)
    _push_all(spreadsheet, Path(file_path), file_format,
              include_formatting=include_formatting)
    return f"Pushed all sheets from {file_path} → {spreadsheet_url}"
```

- [ ] **Step 5: Add include_formatting to pull_rows and push_rows**

In `gssync/mcp_server.py`, replace `pull_rows` with:

```python
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
    include_formatting: bool = True,
) -> str:
    from .storage import read_local, write_local
    client = get_client()
    spreadsheet = open_spreadsheet(client, spreadsheet_url)
    ws = spreadsheet.worksheet(sheet_name)
    all_rows = ws.get_all_values(value_render_option="FORMULA")
    indices = resolve_filter(all_rows, row_numbers, filter_column, filter_value, cell_range)
    if not all_rows:
        header, rows = [], []
    else:
        header = all_rows[0]
        data_rows = all_rows[1:]
        rows = [data_rows[i - 1] for i in indices if 1 <= i <= len(data_rows)]
    path = Path(file_path)
    exists = (path.exists() and path.is_dir()) if file_format == "csv" else path.exists()
    existing = read_local(path, file_format) if exists else {}
    existing[sheet_name] = [header] + rows
    write_local(path, file_format, existing)
    note = ""
    if include_formatting and file_format == "xlsx":
        valid = [i for i in indices if 1 <= i <= len(all_rows) - 1]
        row_map = {0: 0}
        for k, i in enumerate(valid, start=1):
            row_map[i] = k  # 0-based GS row i -> file row k
        sf = read_sheet_formatting(spreadsheet, sheet_name)
        apply_formatting_to_xlsx(path, sheet_name, remap_formatting_rows(sf, row_map))
    elif include_formatting and file_format != "xlsx":
        note = " (formatting skipped: only xlsx supports it)"
    return f"Pulled {len(rows)} row(s) from '{sheet_name}' → {file_path}{note}"
```

Replace `push_rows` with:

```python
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
    include_formatting: bool = True,
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
    note = ""
    if include_formatting and file_format == "xlsx":
        row_map = {i: i for i in indices}  # 0-based row identity (file == GS)
        sf = read_xlsx_formatting(path, sheet_name)
        write_sheet_formatting(spreadsheet, sheet_name,
                               remap_formatting_rows(sf, row_map))
    elif include_formatting and file_format != "xlsx":
        note = " (formatting skipped: only xlsx supports it)"
    return f"Pushed {count} row(s) from {file_path} → '{sheet_name}'{note}"
```

- [ ] **Step 6: Run all MCP tests**

Run: `py -m pytest tests/test_mcp_server.py -v`
Expected: all PASS

- [ ] **Step 7: Run the full suite**

Run: `py -m pytest -v`
Expected: all PASS

- [ ] **Step 8: Smoke test the server import**

Run: `py -c "from gssync.mcp_server import mcp; print('OK')"`
Expected: `OK`

- [ ] **Step 9: Update docs**

In `docs/GSSYNC_MCP.md`, add to the "Синхронизация" section (after the existing `push_all` block) and to the "Построчная работа" section:

```markdown
**Форматирование (для всех pull/push инструментов):**
- `include_formatting` — переносить ли форматирование (шрифт, цвет, заливка, выравнивание, ширина столбцов, числовой формат, границы). По умолчанию `true`. Работает только для `xlsx`; для `csv`/`json` форматирование пропускается.
```

In `README.md`, add a bullet under "Что умеет":

```markdown
- Переносить форматирование (шрифт, цвета, заливку, выравнивание, ширину столбцов, числовые форматы, границы) для xlsx — параметр `include_formatting` (по умолчанию включён)
```

- [ ] **Step 10: Commit**

```bash
git add gssync/mcp_server.py tests/test_mcp_server.py docs/GSSYNC_MCP.md README.md
git commit -m "feat: thread include_formatting through MCP sheet and row tools"
```

---

## Definition of Done

- `py -m pytest -v` is all green (65 prior + new formatting tests).
- `py -c "from gssync.mcp_server import mcp; print('OK')"` prints `OK`.
- `include_formatting` defaults to `True` on all six pull/push tools; csv/json skip formatting with a note in the return string.
- No change to existing value-sync behavior or its tests.
