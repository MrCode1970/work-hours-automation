from datetime import datetime

import gspread
from google.oauth2.service_account import Credentials


def open_spreadsheet(gsheet_id: str, google_json_file: str):
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_file(google_json_file, scopes=scopes)
    client = gspread.authorize(creds)
    return client.open_by_key(gsheet_id)


def month_sheet_name(now: datetime | None = None) -> str:
    """
    Имя листа в формате M.YY (например, 1.26).
    """
    if now is None:
        now = datetime.now()
    return f"{now.month}.{now.strftime('%y')}"


def get_worksheet(spreadsheet, sheet_name: str):
    return spreadsheet.worksheet(sheet_name)
