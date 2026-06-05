from pathlib import Path
from typing import List

import gspread
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow

from .config import CREDENTIALS_FILE, TOKEN_FILE

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]


def get_client() -> gspread.Client:
    creds = None
    if TOKEN_FILE.exists():
        creds = Credentials.from_authorized_user_file(str(TOKEN_FILE), SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(str(CREDENTIALS_FILE), SCOPES)
            creds = flow.run_local_server(port=0)
        TOKEN_FILE.parent.mkdir(parents=True, exist_ok=True)
        TOKEN_FILE.write_text(creds.to_json())
    return gspread.authorize(creds)


def open_spreadsheet(client: gspread.Client, url: str) -> gspread.Spreadsheet:
    return client.open_by_url(url)


def list_sheet_names(spreadsheet: gspread.Spreadsheet) -> List[str]:
    return [ws.title for ws in spreadsheet.worksheets()]


def _normalize(cell: object) -> object:
    if isinstance(cell, (str, int, float, type(None))):
        return cell
    return str(cell)


def read_sheet(spreadsheet: gspread.Spreadsheet, name: str) -> List[List]:
    rows = spreadsheet.worksheet(name).get_all_values(value_render_option="FORMULA")
    return [[_normalize(cell) for cell in row] for row in rows]


def write_sheet(spreadsheet: gspread.Spreadsheet, name: str, data: List[List]) -> None:
    try:
        ws = spreadsheet.worksheet(name)
    except gspread.WorksheetNotFound:
        rows = max(len(data), 1)
        cols = max(max((len(r) for r in data), default=1), 1)
        ws = spreadsheet.add_worksheet(title=name, rows=rows, cols=cols)
    ws.clear()
    if data:
        ws.update(data, value_input_option="USER_ENTERED")


def read_sheet_formatting(spreadsheet: gspread.Spreadsheet, name: str):
    from .formatting import gs_grid_to_sheet_formatting, SheetFormatting
    fields = (
        "sheets(properties(sheetId,title),"
        "data(rowData(values(userEnteredFormat)),"
        "columnMetadata(pixelSize),rowMetadata(pixelSize)))"
    )
    meta = spreadsheet.fetch_sheet_metadata(params={
        "includeGridData": True,
        "ranges": [name],
        "fields": fields,
    })
    for sheet in meta.get("sheets", []):
        if sheet.get("properties", {}).get("title") == name:
            return gs_grid_to_sheet_formatting(sheet)
    return SheetFormatting()


def write_sheet_formatting(spreadsheet: gspread.Spreadsheet, name: str, sf) -> None:
    from .formatting import ir_to_gs_requests
    sheet_id = spreadsheet.worksheet(name).id
    requests = ir_to_gs_requests(sheet_id, sf)
    if requests:
        spreadsheet.batch_update({"requests": requests})
