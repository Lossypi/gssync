"""Backend-neutral formatting intermediate representation (IR) and pure conversions.

No network or file I/O lives here — every function is pure and unit-testable.
Coordinates are 0-based (row, col). Colors are '#RRGGBB'.
"""
from dataclasses import dataclass, field
from typing import Optional

from openpyxl.styles import Font, PatternFill, Alignment, Border, Side


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
