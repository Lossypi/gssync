from pathlib import Path

from fastmcp import FastMCP

from .config import load_config, save_config
from .sheets import get_client, list_sheet_names, open_spreadsheet
from .storage import list_local_sheets as _list_local_sheets
from .sync import pull_all as _pull_all, pull_sheet as _pull_sheet
from .sync import push_all as _push_all, push_sheet as _push_sheet

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


if __name__ == "__main__":
    mcp.run()
