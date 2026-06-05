import pytest
from pathlib import Path
from unittest.mock import MagicMock, patch

from gssync.config import Config
from gssync.mcp_server import (
    get_config,
    get_rows,
    list_google_sheets,
    list_local_sheets,
    pull_all,
    pull_rows,
    pull_sheet,
    push_all,
    push_rows,
    push_sheet,
    set_config,
    update_rows,
)


def test_get_config_returns_formatted_string():
    mock_cfg = Config(
        spreadsheet_url="https://docs.google.com/spreadsheets/d/abc123",
        file_path="C:\\data\\report.xlsx",
        file_format="xlsx",
    )
    with patch("gssync.mcp_server.load_config", return_value=mock_cfg):
        result = get_config()
    assert "https://docs.google.com/spreadsheets/d/abc123" in result
    assert "C:\\data\\report.xlsx" in result
    assert "xlsx" in result


def test_set_config_updates_url():
    mock_cfg = Config(spreadsheet_url="old_url", file_path="old_path", file_format="xlsx")
    with patch("gssync.mcp_server.load_config", return_value=mock_cfg), \
         patch("gssync.mcp_server.save_config") as mock_save:
        result = set_config(spreadsheet_url="new_url")
    mock_save.assert_called_once()
    saved_cfg = mock_save.call_args[0][0]
    assert saved_cfg.spreadsheet_url == "new_url"
    assert saved_cfg.file_path == "old_path"
    assert "new_url" in result


def test_set_config_skips_empty_strings():
    mock_cfg = Config(spreadsheet_url="keep_url", file_path="keep_path", file_format="xlsx")
    with patch("gssync.mcp_server.load_config", return_value=mock_cfg), \
         patch("gssync.mcp_server.save_config") as mock_save:
        set_config()
    saved_cfg = mock_save.call_args[0][0]
    assert saved_cfg.spreadsheet_url == "keep_url"
    assert saved_cfg.file_path == "keep_path"


def test_list_google_sheets_returns_sheet_names():
    mock_spreadsheet = MagicMock()
    with patch("gssync.mcp_server.get_client"), \
         patch("gssync.mcp_server.open_spreadsheet", return_value=mock_spreadsheet), \
         patch("gssync.mcp_server.list_sheet_names", return_value=["Sheet1", "Budget", "Raw"]):
        result = list_google_sheets("https://docs.google.com/spreadsheets/d/abc")
    assert "Sheet1" in result
    assert "Budget" in result
    assert "Raw" in result


def test_list_local_sheets_returns_sheet_names():
    with patch("gssync.mcp_server._list_local_sheets", return_value=["Data", "Summary"]):
        result = list_local_sheets("C:\\data\\report.xlsx")
    assert "Data" in result
    assert "Summary" in result


def test_list_local_sheets_empty_file():
    with patch("gssync.mcp_server._list_local_sheets", return_value=[]):
        result = list_local_sheets("C:\\data\\report.xlsx")
    assert "No sheets" in result


def test_pull_sheet_calls_sync_and_returns_message():
    mock_spreadsheet = MagicMock()
    with patch("gssync.mcp_server.get_client"), \
         patch("gssync.mcp_server.open_spreadsheet", return_value=mock_spreadsheet), \
         patch("gssync.mcp_server._pull_sheet") as mock_pull:
        result = pull_sheet(
            spreadsheet_url="https://docs.google.com/spreadsheets/d/abc",
            sheet_name="Sheet1",
            file_path="C:\\data\\report.xlsx",
        )
    mock_pull.assert_called_once_with(
        mock_spreadsheet, "Sheet1", Path("C:\\data\\report.xlsx"), "xlsx",
        include_formatting=True,
    )
    assert "Sheet1" in result
    assert "C:\\data\\report.xlsx" in result


def test_pull_all_calls_sync_and_returns_count():
    mock_spreadsheet = MagicMock()
    with patch("gssync.mcp_server.get_client"), \
         patch("gssync.mcp_server.open_spreadsheet", return_value=mock_spreadsheet), \
         patch("gssync.mcp_server.list_sheet_names", return_value=["A", "B", "C"]), \
         patch("gssync.mcp_server._pull_all") as mock_pull_all:
        result = pull_all(
            spreadsheet_url="https://docs.google.com/spreadsheets/d/abc",
            file_path="C:\\data\\report.xlsx",
        )
    mock_pull_all.assert_called_once()
    assert "3" in result


def test_push_sheet_calls_sync_and_returns_message():
    mock_spreadsheet = MagicMock()
    with patch("gssync.mcp_server.get_client"), \
         patch("gssync.mcp_server.open_spreadsheet", return_value=mock_spreadsheet), \
         patch("gssync.mcp_server._push_sheet") as mock_push:
        result = push_sheet(
            spreadsheet_url="https://docs.google.com/spreadsheets/d/abc",
            sheet_name="Sheet1",
            file_path="C:\\data\\report.xlsx",
        )
    mock_push.assert_called_once_with(
        mock_spreadsheet, "Sheet1", Path("C:\\data\\report.xlsx"), "xlsx",
        include_formatting=True,
    )
    assert "Sheet1" in result


def test_push_all_calls_sync_and_returns_message():
    mock_spreadsheet = MagicMock()
    with patch("gssync.mcp_server.get_client"), \
         patch("gssync.mcp_server.open_spreadsheet", return_value=mock_spreadsheet), \
         patch("gssync.mcp_server._push_all") as mock_push_all:
        result = push_all(
            spreadsheet_url="https://docs.google.com/spreadsheets/d/abc",
            file_path="C:\\data\\report.xlsx",
        )
    mock_push_all.assert_called_once()
    assert "C:\\data\\report.xlsx" in result


# ── get_rows ──────────────────────────────────────────────────────────────────

def test_get_rows_by_row_numbers():
    mock_ws = MagicMock()
    mock_ws.get_all_values.return_value = [
        ["name", "status"],
        ["Ivan", "active"],
        ["Maria", "done"],
    ]
    mock_ss = MagicMock()
    mock_ss.worksheet.return_value = mock_ws
    with patch("gssync.mcp_server.get_client"), \
         patch("gssync.mcp_server.open_spreadsheet", return_value=mock_ss):
        result = get_rows(
            spreadsheet_url="https://docs.google.com/spreadsheets/d/abc",
            sheet_name="Sheet1",
            row_numbers="1",
        )
    assert "Ivan" in result
    assert "active" in result


def test_get_rows_no_filter_raises():
    mock_ws = MagicMock()
    mock_ws.get_all_values.return_value = [["name"], ["Ivan"]]
    mock_ss = MagicMock()
    mock_ss.worksheet.return_value = mock_ws
    with patch("gssync.mcp_server.get_client"), \
         patch("gssync.mcp_server.open_spreadsheet", return_value=mock_ss):
        with pytest.raises(ValueError, match="Provide exactly one filter"):
            get_rows(
                spreadsheet_url="https://docs.google.com/spreadsheets/d/abc",
                sheet_name="Sheet1",
            )


# ── update_rows ───────────────────────────────────────────────────────────────

def test_update_rows_by_row_numbers():
    mock_ws = MagicMock()
    mock_ws.get_all_values.return_value = [
        ["name", "status"],
        ["Ivan", "active"],
    ]
    mock_ss = MagicMock()
    mock_ss.worksheet.return_value = mock_ws
    with patch("gssync.mcp_server.get_client"), \
         patch("gssync.mcp_server.open_spreadsheet", return_value=mock_ss):
        result = update_rows(
            spreadsheet_url="https://docs.google.com/spreadsheets/d/abc",
            sheet_name="Sheet1",
            rows_data='[{"name": "Ivan", "status": "done"}]',
            row_numbers="1",
        )
    mock_ws.update.assert_called_once_with(
        "A2:B2", [["Ivan", "done"]], value_input_option="USER_ENTERED"
    )
    assert "1" in result


def test_update_rows_broadcast_filter_column():
    mock_ws = MagicMock()
    mock_ws.get_all_values.return_value = [
        ["name", "status"],
        ["Ivan", "active"],
        ["Maria", "active"],
    ]
    mock_ss = MagicMock()
    mock_ss.worksheet.return_value = mock_ws
    with patch("gssync.mcp_server.get_client"), \
         patch("gssync.mcp_server.open_spreadsheet", return_value=mock_ss):
        result = update_rows(
            spreadsheet_url="https://docs.google.com/spreadsheets/d/abc",
            sheet_name="Sheet1",
            rows_data='[{"name": "Ivan", "status": "done"}]',
            filter_column="status",
            filter_value="active",
        )
    assert mock_ws.update.call_count == 2
    assert "2" in result


# ── pull_rows ─────────────────────────────────────────────────────────────────

def test_pull_rows_writes_filtered_rows_to_file(tmp_path):
    mock_ws = MagicMock()
    mock_ws.get_all_values.return_value = [
        ["name", "status"],
        ["Ivan", "active"],
        ["Maria", "done"],
    ]
    mock_ss = MagicMock()
    mock_ss.worksheet.return_value = mock_ws
    xlsx = tmp_path / "data.xlsx"
    with patch("gssync.mcp_server.get_client"), \
         patch("gssync.mcp_server.open_spreadsheet", return_value=mock_ss):
        result = pull_rows(
            spreadsheet_url="https://docs.google.com/spreadsheets/d/abc",
            sheet_name="Sheet1",
            file_path=str(xlsx),
            row_numbers="1",
            include_formatting=False,
        )
    from gssync.storage import read_local
    data = read_local(xlsx, "xlsx")
    assert data["Sheet1"] == [["name", "status"], ["Ivan", "active"]]
    assert "1" in result


# ── push_rows ─────────────────────────────────────────────────────────────────

def test_push_rows_from_local_file(tmp_path):
    from gssync.storage import write_local
    xlsx = tmp_path / "data.xlsx"
    write_local(xlsx, "xlsx", {
        "Sheet1": [["name", "status"], ["Ivan", "active"], ["Maria", "done"]]
    })
    mock_ws = MagicMock()
    mock_ss = MagicMock()
    mock_ss.worksheet.return_value = mock_ws
    with patch("gssync.mcp_server.get_client"), \
         patch("gssync.mcp_server.open_spreadsheet", return_value=mock_ss):
        result = push_rows(
            spreadsheet_url="https://docs.google.com/spreadsheets/d/abc",
            sheet_name="Sheet1",
            file_path=str(xlsx),
            row_numbers="2",
            include_formatting=False,
        )
    mock_ws.update.assert_called_once_with(
        "A3:B3", [["Maria", "done"]], value_input_option="USER_ENTERED"
    )
    assert "1" in result


# ── include_formatting ─────────────────────────────────────────────────────────

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
