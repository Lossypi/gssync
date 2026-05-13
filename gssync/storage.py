import csv
import json
from pathlib import Path
from typing import Dict, List

import openpyxl

SheetData = Dict[str, List[List]]


def read_xlsx(path: Path) -> SheetData:
    wb = openpyxl.load_workbook(path)
    result = {}
    for name in wb.sheetnames:
        ws = wb[name]
        result[name] = [[cell.value for cell in row] for row in ws.iter_rows()]
    return result


def write_xlsx(path: Path, data: SheetData) -> None:
    wb = openpyxl.load_workbook(path) if path.exists() else openpyxl.Workbook()
    if not path.exists() and "Sheet" in wb.sheetnames:
        del wb["Sheet"]
    for name, rows in data.items():
        if name in wb.sheetnames:
            del wb[name]
        ws = wb.create_sheet(title=name)
        for row in rows:
            ws.append([v if v is not None else "" for v in row])
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def read_json(path: Path) -> SheetData:
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def write_json(path: Path, data: SheetData) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def read_csv_dir(path: Path) -> SheetData:
    result = {}
    for csv_file in sorted(path.glob("*.csv")):
        with open(csv_file, newline="", encoding="utf-8") as f:
            result[csv_file.stem] = list(csv.reader(f))
    return result


def write_csv_dir(path: Path, data: SheetData) -> None:
    path.mkdir(parents=True, exist_ok=True)
    for name, rows in data.items():
        with open(path / f"{name}.csv", "w", newline="", encoding="utf-8") as f:
            csv.writer(f).writerows(rows)


def read_local(path: Path, fmt: str) -> SheetData:
    if fmt == "xlsx":
        return read_xlsx(path)
    if fmt == "json":
        return read_json(path)
    if fmt == "csv":
        return read_csv_dir(path)
    raise ValueError(f"Unknown format: {fmt}")


def write_local(path: Path, fmt: str, data: SheetData) -> None:
    if fmt == "xlsx":
        write_xlsx(path, data)
    elif fmt == "json":
        write_json(path, data)
    elif fmt == "csv":
        write_csv_dir(path, data)
    else:
        raise ValueError(f"Unknown format: {fmt}")


def list_local_sheets(path: Path, fmt: str) -> List[str]:
    exists = (path.exists() and path.is_dir()) if fmt == "csv" else path.exists()
    if not exists:
        return []
    return list(read_local(path, fmt).keys())
