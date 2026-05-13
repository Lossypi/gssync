import json
from pathlib import Path
import pytest
from gssync.storage import read_local, write_local, list_local_sheets

SAMPLE = {
    "Sheet1": [["Name", "Age"], ["Alice", "30"], ["Bob", "25"]],
    "Sheet2": [["X", "Y"], ["1", "2"]],
}

def test_xlsx_roundtrip(tmp_path):
    path = tmp_path / "test.xlsx"
    write_local(path, "xlsx", SAMPLE)
    result = read_local(path, "xlsx")
    assert result["Sheet1"] == SAMPLE["Sheet1"]
    assert result["Sheet2"] == SAMPLE["Sheet2"]

def test_xlsx_preserves_existing_sheets(tmp_path):
    path = tmp_path / "test.xlsx"
    write_local(path, "xlsx", SAMPLE)
    write_local(path, "xlsx", {"Sheet3": [["A"], ["1"]]})
    result = read_local(path, "xlsx")
    assert "Sheet1" in result
    assert "Sheet3" in result

def test_json_roundtrip(tmp_path):
    path = tmp_path / "test.json"
    write_local(path, "json", SAMPLE)
    result = read_local(path, "json")
    assert result == SAMPLE

def test_csv_roundtrip(tmp_path):
    path = tmp_path / "csv_dir"
    write_local(path, "csv", SAMPLE)
    result = read_local(path, "csv")
    assert result["Sheet1"] == SAMPLE["Sheet1"]
    assert result["Sheet2"] == SAMPLE["Sheet2"]

def test_list_local_sheets_xlsx(tmp_path):
    path = tmp_path / "test.xlsx"
    write_local(path, "xlsx", SAMPLE)
    names = list_local_sheets(path, "xlsx")
    assert set(names) == {"Sheet1", "Sheet2"}

def test_list_local_sheets_missing_file(tmp_path):
    path = tmp_path / "missing.xlsx"
    assert list_local_sheets(path, "xlsx") == []
