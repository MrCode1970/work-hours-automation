from config import load_config
from sheets_client import open_spreadsheet, month_sheet_name, get_worksheet
from sync_logic import build_changes_sheet
from ylm_portal import download_excel


def main() -> None:
    cfg = load_config()

    # 1. Получаем Excel
    if cfg.get("SKIP_DOWNLOAD"):
        excel_path = cfg["EXCEL_PATH"]
        print(f"⏭️ SKIP_DOWNLOAD=1 — используем локальный Excel: {excel_path}")
    else:
        excel_path = download_excel(
            site_username=cfg["SITE_USERNAME"],
            site_password=cfg["SITE_PASSWORD"],
            excel_path=cfg["EXCEL_PATH"],
            headless=cfg["HEADLESS"],
        )

    # 2. Открываем Google Sheets
    spreadsheet = open_spreadsheet(
        gsheet_id=cfg["GSHEET_ID"],
        google_json_file=cfg["GOOGLE_JSON_FILE"],
    )

    # 3. Получаем основной лист месяца (эталон)
    sheet_name = month_sheet_name()
    worksheet = get_worksheet(spreadsheet, sheet_name)

    # 4. Строим лист "Изменения M.YY"
    build_changes_sheet(
        spreadsheet=spreadsheet,
        base_ws=worksheet,
        sheet_name=sheet_name,
        excel_path=excel_path,
    )

    print("✅ Готово")


if __name__ == "__main__":
    main()
