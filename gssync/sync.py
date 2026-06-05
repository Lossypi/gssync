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
