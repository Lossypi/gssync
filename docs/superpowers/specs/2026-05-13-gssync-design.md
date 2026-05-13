# GSSync — Design Spec

**Date:** 2026-05-13  
**Status:** Approved

## Overview

A TUI application for bidirectional synchronization between a Google Spreadsheet and a local file. Runs in the terminal, uses OAuth for Google authentication, defaults to xlsx format.

## Architecture

```
GSSync/
├── gssync/
│   ├── __main__.py     # entry point: python -m gssync
│   ├── app.py          # Textual TUI application
│   ├── sheets.py       # gspread wrapper (Google Sheets API)
│   ├── storage.py      # local file I/O (xlsx/json/csv)
│   ├── sync.py         # pull/push operation logic
│   └── config.py       # persistence of last URL + file path
├── pyproject.toml
└── requirements.txt
```

**User data directory:** `~/.gssync/`
- `config.json` — last spreadsheet URL + local file path
- `token.json` — OAuth token (saved after first login, auto-refreshed)
- `credentials.json` — OAuth client credentials (user provides once from Google Cloud Console)

**Dependencies:**
- `textual` — TUI framework
- `gspread` — Google Sheets API client
- `google-auth-oauthlib` — OAuth 2.0 flow
- `openpyxl` — xlsx read/write

## First-Run Flow

1. No `~/.gssync/config.json` → show setup form (spreadsheet URL + local file path)
2. No `~/.gssync/token.json` → launch browser for OAuth consent
3. Token saved → main TUI opens
4. Subsequent runs: load config, open main TUI directly (settings changeable via `E`)

## TUI Layout

```
┌─ GSSync ──────────────────────────────────────────────────────┐
│ 📊 My Spreadsheet          📁 C:\data\report.xlsx             │
├───────────────────────┬───────────────────────────────────────┤
│  Google Sheets        │  Local File                           │
│ ─────────────────     │ ─────────────────                     │
│ ▶ Sheet1              │   Sheet1                              │
│   Sheet2              │   Sheet2                              │
│   Summary             │   Summary                             │
│                       │   OldSheet                            │
├───────────────────────┴───────────────────────────────────────┤
│ [R] Refresh  [→] Pull sheet  [←] Push sheet                   │
│ [P] Pull all  [U] Push all  [F] Format  [E] Edit paths  [Q]   │
├───────────────────────────────────────────────────────────────┤
│ ✓ Pulled "Sheet1" successfully                                 │
└───────────────────────────────────────────────────────────────┘
```

**Keyboard bindings:**
| Key | Action |
|-----|--------|
| `↑/↓` | Navigate sheets in active panel |
| `Tab` | Switch active panel (Google ↔ Local) |
| `→` | Pull selected sheet: Google → Local file |
| `←` | Push selected sheet: Local file → Google |
| `P` | Pull All: all Google sheets → local file |
| `U` | Push All: all local sheets → Google |
| `R` | Refresh sheet lists from both sides |
| `F` | Change file format (xlsx / json / csv) |
| `E` | Edit spreadsheet URL and file path |
| `Q` | Quit |

Sheets are matched by name. Sheets that exist on only one side are visible in their respective panel only — no special icon needed.

## Sync Logic

### Pull Sheet
1. Fetch sheet data via `gspread` (returns list of lists)
2. Open local file (create if absent)
3. Find sheet by name → overwrite if exists, create if not
4. Save file

### Pull All
1. Fetch all sheets from Google
2. For each sheet: overwrite or create in local file
3. Does not delete local sheets that are absent in Google

### Push Sheet
1. Read sheet data from local file via `openpyxl`
2. Find sheet in Google by name → if exists: `clear()` + `update()`; if not: `add_worksheet()` then `update()`

### Push All
1. Read all sheets from local file
2. For each sheet: push to Google (same logic as Push Sheet)
3. Does not delete Google sheets absent in the local file

### Sheet Matching
Sheets are matched strictly by name (case-sensitive). No rename detection.

## File Formats

| Format | Storage | Notes |
|--------|---------|-------|
| `xlsx` (default) | Single file, sheets = workbook tabs | via `openpyxl` |
| `json` | Single file: `{"Sheet1": [[...]], ...}` | human-readable structure |
| `csv` | Directory alongside the configured path, one `SheetName.csv` per sheet | loses formatting |

Format is selectable via `F` in the TUI and persisted in `config.json`.

## Error Handling

Errors (no network, file locked, no API access) are displayed in the status bar at the bottom in red. They do not crash the application. The user can retry the operation after resolving the issue.
