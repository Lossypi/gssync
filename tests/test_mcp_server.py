from pathlib import Path
from unittest.mock import MagicMock, patch

from gssync.config import Config
from gssync.mcp_server import get_config, set_config


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
