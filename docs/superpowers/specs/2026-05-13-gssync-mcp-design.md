# GSSync MCP Server Design

## Goal

Add an MCP server to GSSync so Claude can issue sync commands directly from chat without opening the TUI.

## Architecture

A single new file `gssync/mcp_server.py` built with the FastMCP Python SDK. It runs as a stdio subprocess spawned by Claude Desktop. All business logic delegates to existing modules (`sheets.py`, `sync.py`, `storage.py`, `config.py`) — no duplication of sync or auth logic.

Auth reuses the existing OAuth2 flow: `get_client()` in `sheets.py` reads `~/.gssync/token.json` and refreshes automatically. Each MCP tool call opens a fresh gspread client (stateless — no shared session state between calls).

## File Changes

| File | Action |
|------|--------|
| `gssync/mcp_server.py` | Create — FastMCP server with all 7 tools |
| `pyproject.toml` | Add `fastmcp>=2.0` to dependencies |

## Tools

### Sync tools (require `spreadsheet_url`, `file_path`, optional `file_format`)

| Tool | Description |
|------|-------------|
| `pull_sheet` | Pull one named sheet from Google Sheets → local file |
| `pull_all` | Pull all sheets from Google Sheets → local file |
| `push_sheet` | Push one named sheet from local file → Google Sheets |
| `push_all` | Push all local sheets → Google Sheets |

`file_format` defaults to `"xlsx"` when omitted. Valid values: `"xlsx"`, `"json"`, `"csv"`.

### List tools

| Tool | Parameters | Description |
|------|-----------|-------------|
| `list_google_sheets` | `spreadsheet_url` | Returns list of sheet names in the Google spreadsheet |
| `list_local_sheets` | `file_path`, `file_format?` | Returns list of sheet names in the local file |

### Config tools

| Tool | Parameters | Description |
|------|-----------|-------------|
| `get_config` | *(none)* | Returns current saved config (url, path, format) as a readable string |
| `set_config` | `spreadsheet_url?`, `file_path?`, `file_format?` | Updates one or more config fields, returns confirmation |

## Return Values

All tools return plain strings describing the outcome. Sync tools include the sheet name(s) and file path. Example:

- `pull_sheet` → `"Pulled 'Sheet1' from <url> → C:\data\report.xlsx"`
- `pull_all` → `"Pulled 3 sheets from <url> → C:\data\report.xlsx"`
- `list_google_sheets` → `"Sheets: Sheet1, Sheet2, Budget"`

Errors propagate as raised exceptions (FastMCP converts them to error responses automatically).

## Entry Point

```python
if __name__ == "__main__":
    mcp.run()
```

Run via: `py -m gssync.mcp_server`

## Claude Desktop Registration

User adds to `%APPDATA%\Claude\claude_desktop_config.json`:

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

## Dependencies

Add to `pyproject.toml` `[project]` dependencies:

```
"fastmcp>=2.0",
```

## Error Handling

- Missing `credentials.json` or expired token: exception propagates with the original message from `get_client()`
- Unknown sheet name: exception from `gspread.WorksheetNotFound` propagates
- Unknown format: `ValueError` from `storage.py` propagates
- No validation beyond what existing modules already do
