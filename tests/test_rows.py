import pytest
from pathlib import Path
from unittest.mock import MagicMock

from gssync.rows import (
    _col_index_to_letter,
    _col_letter_to_index,
    filter_rows_by_column,
    parse_cell_range,
    parse_row_numbers,
    read_range_from_sheet,
    read_rows_from_local,
    read_rows_from_sheet,
    resolve_filter,
    rows_data_to_lists,
    write_range_to_sheet,
    write_rows_to_sheet,
)


# ── parse_row_numbers ────────────────────────────────────────────────────────

def test_parse_row_numbers_single():
    assert parse_row_numbers("5") == [5]

def test_parse_row_numbers_list():
    assert parse_row_numbers("2,5,10") == [2, 5, 10]

def test_parse_row_numbers_range():
    assert parse_row_numbers("2:5") == [2, 3, 4, 5]

def test_parse_row_numbers_single_item_list():
    assert parse_row_numbers("3,") == [3]


# ── _col_letter_to_index / _col_index_to_letter ──────────────────────────────

def test_col_letter_to_index():
    assert _col_letter_to_index("A") == 1
    assert _col_letter_to_index("Z") == 26
    assert _col_letter_to_index("AA") == 27

def test_col_index_to_letter():
    assert _col_index_to_letter(1) == "A"
    assert _col_index_to_letter(26) == "Z"
    assert _col_index_to_letter(27) == "AA"

def test_col_roundtrip():
    for i in range(1, 30):
        assert _col_letter_to_index(_col_index_to_letter(i)) == i


# ── parse_cell_range ─────────────────────────────────────────────────────────

def test_parse_cell_range_simple():
    assert parse_cell_range("A2:D10") == (2, 10, 1, 4)

def test_parse_cell_range_lowercase():
    assert parse_cell_range("a2:d10") == (2, 10, 1, 4)

def test_parse_cell_range_invalid():
    with pytest.raises(ValueError, match="Invalid cell range"):
        parse_cell_range("bad")


# ── filter_rows_by_column ────────────────────────────────────────────────────

def test_filter_rows_by_column_matches():
    header = ["name", "status"]
    rows = [["Ivan", "active"], ["Maria", "done"], ["Petr", "active"]]
    assert filter_rows_by_column(rows, header, "status", "active") == [1, 3]

def test_filter_rows_by_column_no_match():
    header = ["name", "status"]
    rows = [["Ivan", "active"]]
    assert filter_rows_by_column(rows, header, "status", "done") == []

def test_filter_rows_by_column_missing_column():
    with pytest.raises(ValueError, match="Column 'missing'"):
        filter_rows_by_column([], ["name"], "missing", "x")

def test_filter_rows_by_column_short_row_skipped():
    header = ["name", "status", "amount"]
    rows = [["Ivan", "active"], ["Maria", "active", "100"]]  # first row lacks 'amount'
    # Both should match on 'status' column (index 1), which is present in both
    result = filter_rows_by_column(rows, header, "status", "active")
    assert result == [1, 2]


# ── resolve_filter ───────────────────────────────────────────────────────────

def test_resolve_filter_row_numbers():
    assert resolve_filter([], row_numbers="2:4") == [2, 3, 4]

def test_resolve_filter_column():
    all_rows = [["name", "status"], ["Ivan", "active"], ["Maria", "done"]]
    assert resolve_filter(all_rows, filter_column="status", filter_value="active") == [1]

def test_resolve_filter_cell_range():
    assert resolve_filter([], cell_range="A3:D5") == [2, 3, 4]

def test_resolve_filter_no_filter():
    with pytest.raises(ValueError, match="Provide exactly one filter"):
        resolve_filter([])

def test_resolve_filter_multiple_filters():
    with pytest.raises(ValueError, match="Provide exactly one filter"):
        resolve_filter([], row_numbers="1", cell_range="A2:B3")

def test_resolve_filter_cell_range_row1_raises():
    with pytest.raises(ValueError, match="row 2 or later"):
        resolve_filter([], cell_range="A1:D5")

def test_resolve_filter_cell_range_single_row():
    assert resolve_filter([], cell_range="A2:D2") == [1]


# ── rows_data_to_lists ───────────────────────────────────────────────────────

def test_rows_data_to_lists_basic():
    header = ["name", "status"]
    result = rows_data_to_lists('[{"name": "Ivan", "status": "done"}]', header)
    assert result == [["Ivan", "done"]]

def test_rows_data_to_lists_missing_key():
    header = ["name", "status"]
    result = rows_data_to_lists('[{"name": "Ivan"}]', header)
    assert result == [["Ivan", ""]]

def test_rows_data_to_lists_invalid_json():
    with pytest.raises(ValueError, match="valid JSON"):
        rows_data_to_lists("not json", ["name"])

def test_rows_data_to_lists_not_list():
    with pytest.raises(ValueError, match="JSON array"):
        rows_data_to_lists('{"name": "Ivan"}', ["name"])


# ── read_rows_from_sheet ─────────────────────────────────────────────────────

def test_read_rows_from_sheet_by_indices():
    mock_ws = MagicMock()
    mock_ws.get_all_values.return_value = [
        ["name", "status"],
        ["Ivan", "active"],
        ["Maria", "done"],
        ["Petr", "active"],
    ]
    mock_ss = MagicMock()
    mock_ss.worksheet.return_value = mock_ws
    header, rows = read_rows_from_sheet(mock_ss, "Sheet1", [1, 3])
    assert header == ["name", "status"]
    assert rows == [["Ivan", "active"], ["Petr", "active"]]

def test_read_rows_from_sheet_empty():
    mock_ws = MagicMock()
    mock_ws.get_all_values.return_value = []
    mock_ss = MagicMock()
    mock_ss.worksheet.return_value = mock_ws
    header, rows = read_rows_from_sheet(mock_ss, "Sheet1", [1])
    assert header == []
    assert rows == []

def test_read_rows_from_sheet_out_of_bounds_skipped():
    mock_ws = MagicMock()
    mock_ws.get_all_values.return_value = [["name"], ["Ivan"]]
    mock_ss = MagicMock()
    mock_ss.worksheet.return_value = mock_ws
    header, rows = read_rows_from_sheet(mock_ss, "Sheet1", [1, 99])
    assert rows == [["Ivan"]]


# ── read_range_from_sheet ────────────────────────────────────────────────────

def test_read_range_from_sheet_returns_header_and_rows():
    mock_ws = MagicMock()
    mock_ws.row_values.return_value = ["name", "status"]
    mock_ws.get.return_value = [["Ivan", "active"], ["Maria", "done"]]
    mock_ss = MagicMock()
    mock_ss.worksheet.return_value = mock_ws
    header, rows = read_range_from_sheet(mock_ss, "Sheet1", "A2:B3")
    mock_ws.row_values.assert_called_once_with(1)
    mock_ws.get.assert_called_once_with("A2:B3", value_render_option="FORMULA")
    assert header == ["name", "status"]
    assert rows == [["Ivan", "active"], ["Maria", "done"]]


# ── write_rows_to_sheet ──────────────────────────────────────────────────────

def test_write_rows_to_sheet_calls_update():
    mock_ws = MagicMock()
    mock_ss = MagicMock()
    mock_ss.worksheet.return_value = mock_ws
    count = write_rows_to_sheet(mock_ss, "Sheet1", [2], [["Ivan", "done"]])
    assert count == 1
    mock_ws.update.assert_called_once_with(
        "A3:B3", [["Ivan", "done"]], value_input_option="USER_ENTERED"
    )

def test_write_rows_to_sheet_mismatched_lengths_raises():
    mock_ss = MagicMock()
    with pytest.raises(ValueError, match="same length"):
        write_rows_to_sheet(mock_ss, "Sheet1", [1, 2], [["only one row"]])

def test_write_rows_to_sheet_skips_empty_row():
    mock_ws = MagicMock()
    mock_ss = MagicMock()
    mock_ss.worksheet.return_value = mock_ws
    count = write_rows_to_sheet(mock_ss, "Sheet1", [1, 2], [[], ["Ivan", "done"]])
    assert count == 1  # empty row skipped
    mock_ws.update.assert_called_once_with(
        "A3:B3", [["Ivan", "done"]], value_input_option="USER_ENTERED"
    )


# ── write_range_to_sheet ─────────────────────────────────────────────────────

def test_write_range_to_sheet_calls_update():
    mock_ws = MagicMock()
    mock_ss = MagicMock()
    mock_ss.worksheet.return_value = mock_ws
    count = write_range_to_sheet(mock_ss, "Sheet1", "B2:C3", [["x", "y"], ["a", "b"]])
    assert count == 2
    mock_ws.update.assert_called_once_with(
        "B2:C3", [["x", "y"], ["a", "b"]], value_input_option="USER_ENTERED"
    )


# ── read_rows_from_local ─────────────────────────────────────────────────────

def test_read_rows_from_local_basic(tmp_path):
    from gssync.storage import write_local
    xlsx = tmp_path / "data.xlsx"
    write_local(xlsx, "xlsx", {"Sheet1": [["name", "status"], ["Ivan", "active"], ["Maria", "done"]]})
    header, rows = read_rows_from_local(xlsx, "xlsx", "Sheet1", [2])
    assert header == ["name", "status"]
    assert rows == [["Maria", "done"]]

def test_read_rows_from_local_missing_sheet(tmp_path):
    from gssync.storage import write_local
    xlsx = tmp_path / "data.xlsx"
    write_local(xlsx, "xlsx", {"Sheet1": [["name"]]})
    with pytest.raises(ValueError, match="Sheet 'Missing'"):
        read_rows_from_local(xlsx, "xlsx", "Missing", [1])
