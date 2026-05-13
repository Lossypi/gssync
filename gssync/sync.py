from pathlib import Path

import gspread

from .sheets import list_sheet_names, read_sheet, write_sheet
from .storage import list_local_sheets, read_local, write_local


def _load_existing(path: Path, fmt: str) -> dict:
    exists = (path.exists() and path.is_dir()) if fmt == "csv" else path.exists()
    return read_local(path, fmt) if exists else {}


def pull_sheet(spreadsheet: gspread.Spreadsheet, sheet_name: str, file_path: Path, fmt: str) -> None:
    data = read_sheet(spreadsheet, sheet_name)
    existing = _load_existing(file_path, fmt)
    existing[sheet_name] = data
    write_local(file_path, fmt, existing)


def pull_all(spreadsheet: gspread.Spreadsheet, file_path: Path, fmt: str) -> None:
    names = list_sheet_names(spreadsheet)
    existing = _load_existing(file_path, fmt)
    for name in names:
        existing[name] = read_sheet(spreadsheet, name)
    write_local(file_path, fmt, existing)


def push_sheet(spreadsheet: gspread.Spreadsheet, sheet_name: str, file_path: Path, fmt: str) -> None:
    local_data = read_local(file_path, fmt)
    write_sheet(spreadsheet, sheet_name, local_data[sheet_name])


def push_all(spreadsheet: gspread.Spreadsheet, file_path: Path, fmt: str) -> None:
    local_data = read_local(file_path, fmt)
    for name, rows in local_data.items():
        write_sheet(spreadsheet, name, rows)
