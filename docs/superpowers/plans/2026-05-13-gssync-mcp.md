# GSSync MCP Server Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a FastMCP stdio server to GSSync exposing 7 tools for syncing, listing, and configuring Google Sheets from Claude chat.

**Architecture:** Single new file `gssync/mcp_server.py` wraps the existing `sync.py`, `sheets.py`, `storage.py`, and `config.py` modules via FastMCP tool decorators. No business logic is duplicated — the MCP layer is thin wiring. Claude Desktop spawns the server as a stdio subprocess via `py -m gssync.mcp_server`.

**Tech Stack:** Python 3.10+, FastMCP 2.x, gspread, openpyxl, google-auth-oauthlib (all pre-existing)

---

## File Map

| File | Action | Responsibility |
|------|--------|----------------|
| `pyproject.toml` | Modify | Add `fastmcp>=2.0` dependency |
| `gssync/mcp_server.py` | Create | All 7 FastMCP tools + `mcp.run()` entry point |
| `tests/test_mcp_server.py` | Create | Unit tests for all 7 tools (mocking external deps) |

---

### Task 1: Add FastMCP dependency

**Files:**
- Modify: `pyproject.toml`

- [ ] **Step 1: Add fastmcp to pyproject.toml**

Replace the `dependencies` list in `pyproject.toml`:

```toml
[project]
name = "gssync"
version = "0.1.0"
requires-python = ">=3.10"
dependencies = [
    "gspread>=6.0",
    "google-auth-oauthlib>=1.2",
    "openpyxl>=3.1",
    "textual>=0.60",
    "fastmcp>=2.0",
]
```

- [ ] **Step 2: Install the new dependency**

Run: `py -m pip install "fastmcp>=2.0"`

Expected: output ends with `Successfully installed fastmcp-...` or `Requirement already satisfied`

- [ ] **Step 3: Verify import works**

Run: `py -c "import fastmcp; print(fastmcp.__version__)"`

Expected: prints a version string like `2.x.x`

- [ ] **Step 4: Commit**

```bash
git add pyproject.toml
git commit -m "feat: add fastmcp dependency"
```

---

### Task 2: Config tools (get_config, set_config)

**Files:**
- Create: `tests/test_mcp_server.py`
- Create: `gssync/mcp_server.py`

- [ ] **Step 1: Write failing tests for get_config and set_config**

Create `tests/test_mcp_server.py`:

```python
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
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `py -m pytest tests/test_mcp_server.py -v`

Expected: `ImportError` — `gssync.mcp_server` does not exist yet

- [ ] **Step 3: Create gssync/mcp_server.py with config tools**

Create `gssync/mcp_server.py`:

```python
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
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `py -m pytest tests/test_mcp_server.py -v`

Expected: 3 tests PASS

- [ ] **Step 5: Commit**

```bash
git add gssync/mcp_server.py tests/test_mcp_server.py
git commit -m "feat: add MCP server with config tools"
```

---

### Task 3: List tools (list_google_sheets, list_local_sheets)

**Files:**
- Modify: `tests/test_mcp_server.py`
- Modify: `gssync/mcp_server.py`

- [ ] **Step 1: Update imports and add failing tests in tests/test_mcp_server.py**

Replace line 4 (`from gssync.mcp_server import get_config, set_config`) with:

```python
from gssync.mcp_server import get_config, list_google_sheets, list_local_sheets, set_config
```

Then append these test functions to the end of `tests/test_mcp_server.py`:

```python
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
```

- [ ] **Step 2: Run tests to verify new tests fail**

Run: `py -m pytest tests/test_mcp_server.py -v`

Expected: 3 existing PASS, 3 new FAIL with `ImportError` for `list_google_sheets`, `list_local_sheets`

- [ ] **Step 3: Add list tools to gssync/mcp_server.py**

Insert these two functions in `gssync/mcp_server.py` before the `if __name__ == "__main__":` line:

```python
@mcp.tool()
def list_google_sheets(spreadsheet_url: str) -> str:
    client = get_client()
    spreadsheet = open_spreadsheet(client, spreadsheet_url)
    names = list_sheet_names(spreadsheet)
    return f"Sheets: {', '.join(names)}"


@mcp.tool()
def list_local_sheets(file_path: str, file_format: str = "xlsx") -> str:
    names = _list_local_sheets(Path(file_path), file_format)
    if not names:
        return f"No sheets found in {file_path}"
    return f"Sheets: {', '.join(names)}"
```

- [ ] **Step 4: Run all tests**

Run: `py -m pytest tests/test_mcp_server.py -v`

Expected: 6 tests PASS

- [ ] **Step 5: Commit**

```bash
git add gssync/mcp_server.py tests/test_mcp_server.py
git commit -m "feat: add MCP list tools"
```

---

### Task 4: Sync tools (pull_sheet, pull_all, push_sheet, push_all)

**Files:**
- Modify: `tests/test_mcp_server.py`
- Modify: `gssync/mcp_server.py`

- [ ] **Step 1: Update imports and add failing tests in tests/test_mcp_server.py**

Replace the import line at the top of `tests/test_mcp_server.py` with:

```python
from gssync.mcp_server import (
    get_config,
    list_google_sheets,
    list_local_sheets,
    pull_all,
    pull_sheet,
    push_all,
    push_sheet,
    set_config,
)
```

Then append these test functions to the end of `tests/test_mcp_server.py`:

```python
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
        mock_spreadsheet, "Sheet1", Path("C:\\data\\report.xlsx"), "xlsx"
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
        mock_spreadsheet, "Sheet1", Path("C:\\data\\report.xlsx"), "xlsx"
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
```

- [ ] **Step 2: Run tests to verify new tests fail**

Run: `py -m pytest tests/test_mcp_server.py -v`

Expected: 6 existing PASS, 4 new FAIL with `ImportError` for `pull_sheet`, `pull_all`, `push_sheet`, `push_all`

- [ ] **Step 3: Add sync tools to gssync/mcp_server.py**

Insert these four functions in `gssync/mcp_server.py` before the `if __name__ == "__main__":` line:

```python
@mcp.tool()
def pull_sheet(
    spreadsheet_url: str,
    sheet_name: str,
    file_path: str,
    file_format: str = "xlsx",
) -> str:
    client = get_client()
    spreadsheet = open_spreadsheet(client, spreadsheet_url)
    _pull_sheet(spreadsheet, sheet_name, Path(file_path), file_format)
    return f"Pulled '{sheet_name}' from {spreadsheet_url} → {file_path}"


@mcp.tool()
def pull_all(
    spreadsheet_url: str,
    file_path: str,
    file_format: str = "xlsx",
) -> str:
    client = get_client()
    spreadsheet = open_spreadsheet(client, spreadsheet_url)
    names = list_sheet_names(spreadsheet)
    _pull_all(spreadsheet, Path(file_path), file_format)
    return f"Pulled {len(names)} sheets from {spreadsheet_url} → {file_path}"


@mcp.tool()
def push_sheet(
    spreadsheet_url: str,
    sheet_name: str,
    file_path: str,
    file_format: str = "xlsx",
) -> str:
    client = get_client()
    spreadsheet = open_spreadsheet(client, spreadsheet_url)
    _push_sheet(spreadsheet, sheet_name, Path(file_path), file_format)
    return f"Pushed '{sheet_name}' from {file_path} → {spreadsheet_url}"


@mcp.tool()
def push_all(
    spreadsheet_url: str,
    file_path: str,
    file_format: str = "xlsx",
) -> str:
    client = get_client()
    spreadsheet = open_spreadsheet(client, spreadsheet_url)
    _push_all(spreadsheet, Path(file_path), file_format)
    return f"Pushed all sheets from {file_path} → {spreadsheet_url}"
```

- [ ] **Step 4: Run all MCP tests**

Run: `py -m pytest tests/test_mcp_server.py -v`

Expected: 10 tests PASS

- [ ] **Step 5: Run full test suite to check for regressions**

Run: `py -m pytest -v`

Expected: All tests PASS (existing tests in `tests/test_config.py`, `tests/test_storage.py`, `tests/test_sync.py` must still pass)

- [ ] **Step 6: Smoke-test the server module loads cleanly**

Run: `py -c "from gssync.mcp_server import mcp; print('Server ready:', mcp.name)"`

Expected: `Server ready: GSSync`

- [ ] **Step 7: Commit**

```bash
git add gssync/mcp_server.py tests/test_mcp_server.py
git commit -m "feat: complete MCP server with sync and list tools"
```

---

## Post-Implementation: Register with Claude Desktop

After the implementation is complete, add the following to `%APPDATA%\Claude\claude_desktop_config.json` (create the file if it doesn't exist):

```json
{
  "mcpServers": {
    "gssync": {
      "command": "py",
      "args": ["-m", "gssync.mcp_server"],
      "cwd": "C:\\Users\\koltr\\GSSync"
    }
  }
}
```

Then restart Claude Desktop. The 7 GSSync tools will appear in the Tools list.
