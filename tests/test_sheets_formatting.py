from unittest.mock import MagicMock

from gssync.formatting import SheetFormatting, CellFormat
from gssync.sheets import read_sheet_formatting, write_sheet_formatting


def test_read_sheet_formatting_parses_metadata():
    mock_ss = MagicMock()
    mock_ss.fetch_sheet_metadata.return_value = {
        "sheets": [{
            "properties": {"sheetId": 7, "title": "Sheet1"},
            "data": [{
                "rowData": [{"values": [{"userEnteredFormat": {"textFormat": {"bold": True}}}]}],
                "columnMetadata": [{"pixelSize": 120}],
                "rowMetadata": [{"pixelSize": 21}],
            }],
        }]
    }
    sf = read_sheet_formatting(mock_ss, "Sheet1")
    assert sf.cells[(0, 0)].bold is True
    assert sf.column_widths == {0: 120}
    mock_ss.fetch_sheet_metadata.assert_called_once()


def test_read_sheet_formatting_missing_sheet_returns_empty():
    mock_ss = MagicMock()
    mock_ss.fetch_sheet_metadata.return_value = {"sheets": []}
    sf = read_sheet_formatting(mock_ss, "Nope")
    assert sf.cells == {}


def test_write_sheet_formatting_calls_batch_update():
    mock_ws = MagicMock()
    mock_ws.id = 7
    mock_ss = MagicMock()
    mock_ss.worksheet.return_value = mock_ws
    sf = SheetFormatting(cells={(0, 0): CellFormat(bold=True)})
    write_sheet_formatting(mock_ss, "Sheet1", sf)
    mock_ss.batch_update.assert_called_once()
    body = mock_ss.batch_update.call_args[0][0]
    assert "requests" in body
    assert body["requests"][0]["repeatCell"]["range"]["sheetId"] == 7


def test_write_sheet_formatting_no_requests_skips_call():
    mock_ws = MagicMock()
    mock_ws.id = 7
    mock_ss = MagicMock()
    mock_ss.worksheet.return_value = mock_ws
    write_sheet_formatting(mock_ss, "Sheet1", SheetFormatting())
    mock_ss.batch_update.assert_not_called()
