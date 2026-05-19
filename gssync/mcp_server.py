import json
from pathlib import Path

from fastmcp import FastMCP

from .config import load_config, save_config
from .sheets import get_client, list_sheet_names, open_spreadsheet
from .storage import list_local_sheets as _list_local_sheets
from .sync import pull_all as _pull_all, pull_sheet as _pull_sheet
from .sync import push_all as _push_all, push_sheet as _push_sheet
from .rows import (
    read_range_from_sheet,
    read_rows_from_local,
    read_rows_from_sheet,
    resolve_filter,
    rows_data_to_lists,
    write_range_to_sheet,
    write_rows_to_sheet,
)

mcp = FastMCP("GSSync")


@mcp.tool()
def get_config() -> str:
    cfg = load_config()
    return (
        f"spreadsheet_url: {cfg.spreadsheet_url}\n"
        f"file_path: {cfg.file_path}\n"
        f"file_format: {cfg.file_format}"
    )


@mcp.tool()
def set_config(
    spreadsheet_url: str = "",
    file_path: str = "",
    file_format: str = "",
) -> str:
    cfg = load_config()
    if spreadsheet_url:
        cfg.spreadsheet_url = spreadsheet_url
    if file_path:
        cfg.file_path = file_path
    if file_format:
        cfg.file_format = file_format
    save_config(cfg)
    return (
        f"Config updated: spreadsheet_url={cfg.spreadsheet_url}, "
        f"file_path={cfg.file_path}, file_format={cfg.file_format}"
    )


@mcp.tool()
def list_google_sheets(spreadsheet_url: str) -> str:
    client = get_client()
    spreadsheet = open_spreadsheet(client, spreadsheet_url)
    names = list_sheet_names(spreadsheet)
    return f"Sheets: {', '.join(names)}"


@mcp.tool()
def list_local_sheets(file_path: str, file_format: str = "xlsx") -> str:
    names = _list_local_sheets(Path(file_path), file_format)
    if not names:
        return f"No sheets found in {file_path}"
    return f"Sheets: {', '.join(names)}"


@mcp.tool()
def pull_sheet(
    spreadsheet_url: str,
    sheet_name: str,
    file_path: str,
    file_format: str = "xlsx",
) -> str:
    client = get_client()
    spreadsheet = open_spreadsheet(client, spreadsheet_url)
    _pull_sheet(spreadsheet, sheet_name, Path(file_path), file_format)
    return f"Pulled '{sheet_name}' from {spreadsheet_url} → {file_path}"


@mcp.tool()
def pull_all(
    spreadsheet_url: str,
    file_path: str,
    file_format: str = "xlsx",
) -> str:
    client = get_client()
    spreadsheet = open_spreadsheet(client, spreadsheet_url)
    names = list_sheet_names(spreadsheet)
    _pull_all(spreadsheet, Path(file_path), file_format)
    return f"Pulled {len(names)} sheets from {spreadsheet_url} → {file_path}"


@mcp.tool()
def push_sheet(
    spreadsheet_url: str,
    sheet_name: str,
    file_path: str,
    file_format: str = "xlsx",
) -> str:
    client = get_client()
    spreadsheet = open_spreadsheet(client, spreadsheet_url)
    _push_sheet(spreadsheet, sheet_name, Path(file_path), file_format)
    return f"Pushed '{sheet_name}' from {file_path} → {spreadsheet_url}"


@mcp.tool()
def push_all(
    spreadsheet_url: str,
    file_path: str,
    file_format: str = "xlsx",
) -> str:
    client = get_client()
    spreadsheet = open_spreadsheet(client, spreadsheet_url)
    _push_all(spreadsheet, Path(file_path), file_format)
    return f"Pushed all sheets from {file_path} → {spreadsheet_url}"


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
        if not all_rows:
            header, rows = [], []
        else:
            header = all_rows[0]
            data_rows = all_rows[1:]
            rows = [data_rows[i - 1] for i in indices if 1 <= i <= len(data_rows)]
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
    client = get_client()
    spreadsheet = open_spreadsheet(client, spreadsheet_url)
    ws = spreadsheet.worksheet(sheet_name)
    if cell_range and not row_numbers and not filter_column:
        # Direct range write — rows_data is a JSON array of arrays
        data = json.loads(rows_data)
        count = write_range_to_sheet(spreadsheet, sheet_name, cell_range, data)
        return f"Updated {count} row(s) in '{sheet_name}' at {cell_range}"
    all_rows = ws.get_all_values(value_render_option="FORMULA")
    header = all_rows[0] if all_rows else []
    indices = resolve_filter(all_rows, row_numbers, filter_column, filter_value, cell_range)
    new_rows = rows_data_to_lists(rows_data, header)
    if filter_column:
        # Broadcast: apply first entry to all matching rows
        broadcast = new_rows[0] if new_rows else []
        write_data = [list(broadcast) for _ in range(len(indices))]
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


if __name__ == "__main__":
    mcp.run()
