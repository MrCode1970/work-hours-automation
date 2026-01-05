from config import load_config
from sheets_client import open_spreadsheet, month_sheet_name, get_worksheet
from sync_logic import sync_from_excel
from ylm_portal import download_excel


def main() -> None:
    cfg = load_config()

    excel_path = download_excel(
        site_username=cfg["SITE_USERNAME"],
        site_password=cfg["SITE_PASSWORD"],
        excel_path=cfg["EXCEL_PATH"],
        headless=cfg["HEADLESS"],
    )

    spreadsheet = open_spreadsheet(
        gsheet_id=cfg["GSHEET_ID"],
        google_json_file=cfg["GOOGLE_JSON_FILE"],
    )

    sheet_name = month_sheet_name()
    worksheet = get_worksheet(spreadsheet, sheet_name)

    sync_from_excel(worksheet, excel_path)
    print("✅ Готово")


if __name__ == "__main__":
    main()
