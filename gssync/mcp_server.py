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


if __name__ == "__main__":
    mcp.run()
