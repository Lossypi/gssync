from pathlib import Path
from unittest.mock import MagicMock, patch
import pytest
from gssync.sync import pull_sheet, pull_all, push_sheet, push_all


def make_spreadsheet(sheets: dict) -> MagicMock:
    sp = MagicMock()
    sp.worksheets.return_value = [MagicMock(title=k) for k in sheets]
    sp.worksheet.side_effect = lambda name: MagicMock(get_all_values=lambda: sheets[name])
    return sp


def test_pull_sheet_creates_local_file(tmp_path):
    sp = make_spreadsheet({"Sheet1": [["A", "B"], ["1", "2"]]})
    path = tmp_path / "out.xlsx"
    with patch("gssync.sync.read_sheet", return_value=[["A", "B"], ["1", "2"]]):
        pull_sheet(sp, "Sheet1", path, "xlsx")
    from gssync.storage import read_local
    assert read_local(path, "xlsx")["Sheet1"] == [["A", "B"], ["1", "2"]]


def test_pull_sheet_preserves_other_sheets(tmp_path):
    from gssync.storage import write_local
    path = tmp_path / "out.xlsx"
    write_local(path, "xlsx", {"Existing": [["X"]]})
    with patch("gssync.sync.read_sheet", return_value=[["New"]]):
        pull_sheet(MagicMock(), "NewSheet", path, "xlsx")
    from gssync.storage import read_local
    result = read_local(path, "xlsx")
    assert "Existing" in result
    assert "NewSheet" in result


def test_pull_all_writes_all_sheets(tmp_path):
    path = tmp_path / "out.xlsx"
    sheet_data = {"S1": [["a"]], "S2": [["b"]]}
    with patch("gssync.sync.list_sheet_names", return_value=["S1", "S2"]), \
         patch("gssync.sync.read_sheet", side_effect=lambda sp, n: sheet_data[n]):
        pull_all(MagicMock(), path, "xlsx")
    from gssync.storage import read_local
    result = read_local(path, "xlsx")
    assert result["S1"] == [["a"]]
    assert result["S2"] == [["b"]]


def test_push_sheet_calls_write_sheet(tmp_path):
    from gssync.storage import write_local
    path = tmp_path / "data.xlsx"
    write_local(path, "xlsx", {"Sheet1": [["A", "1"]]})
    with patch("gssync.sync.write_sheet") as mock_write:
        push_sheet(MagicMock(), "Sheet1", path, "xlsx")
    mock_write.assert_called_once()
    assert mock_write.call_args[0][1] == "Sheet1"
    assert mock_write.call_args[0][2] == [["A", "1"]]


def test_push_all_pushes_each_sheet(tmp_path):
    from gssync.storage import write_local
    path = tmp_path / "data.xlsx"
    write_local(path, "xlsx", {"S1": [["a"]], "S2": [["b"]]})
    with patch("gssync.sync.write_sheet") as mock_write:
        push_all(MagicMock(), path, "xlsx")
    assert mock_write.call_count == 2


# ── include_formatting wiring ──────────────────────────────────────────────────

def test_pull_sheet_applies_formatting_for_xlsx(tmp_path):
    mock_ss = MagicMock()
    xlsx = tmp_path / "data.xlsx"
    with patch("gssync.sync.read_sheet", return_value=[["a", "b"], ["1", "2"]]), \
         patch("gssync.sheets.read_sheet_formatting") as mock_read_fmt, \
         patch("gssync.storage.apply_formatting_to_xlsx") as mock_apply:
        pull_sheet(mock_ss, "Sheet1", xlsx, "xlsx", include_formatting=True)
    mock_read_fmt.assert_called_once_with(mock_ss, "Sheet1")
    mock_apply.assert_called_once()


def test_pull_sheet_skips_formatting_for_csv(tmp_path):
    mock_ss = MagicMock()
    csv_dir = tmp_path / "out"
    with patch("gssync.sync.read_sheet", return_value=[["a"]]), \
         patch("gssync.sheets.read_sheet_formatting") as mock_read_fmt:
        pull_sheet(mock_ss, "Sheet1", csv_dir, "csv", include_formatting=True)
    mock_read_fmt.assert_not_called()


def test_pull_sheet_skips_formatting_when_flag_off(tmp_path):
    mock_ss = MagicMock()
    xlsx = tmp_path / "data.xlsx"
    with patch("gssync.sync.read_sheet", return_value=[["a"]]), \
         patch("gssync.sheets.read_sheet_formatting") as mock_read_fmt:
        pull_sheet(mock_ss, "Sheet1", xlsx, "xlsx", include_formatting=False)
    mock_read_fmt.assert_not_called()


def test_push_sheet_writes_formatting_for_xlsx(tmp_path):
    from gssync.storage import write_local
    mock_ss = MagicMock()
    xlsx = tmp_path / "data.xlsx"
    write_local(xlsx, "xlsx", {"Sheet1": [["a", "b"]]})
    with patch("gssync.sheets.write_sheet_formatting") as mock_write_fmt, \
         patch("gssync.storage.read_xlsx_formatting", return_value="SF"):
        push_sheet(mock_ss, "Sheet1", xlsx, "xlsx", include_formatting=True)
    mock_write_fmt.assert_called_once_with(mock_ss, "Sheet1", "SF")
