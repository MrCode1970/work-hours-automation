import os
import shutil
from datetime import datetime

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
        excel_path = cfg["EXCEL_PATH"]
        temp_path = f"{excel_path}.new"
        download_excel(
            site_username=cfg["SITE_USERNAME"],
            site_password=cfg["SITE_PASSWORD"],
            excel_path=temp_path,
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
    changes_found = build_changes_sheet(
        spreadsheet=spreadsheet,
        base_ws=worksheet,
        sheet_name=sheet_name,
        excel_path=temp_path if not cfg.get("SKIP_DOWNLOAD") else excel_path,
    )

    if not cfg.get("SKIP_DOWNLOAD"):
        if changes_found and os.path.exists(excel_path):
            history_dir = "history"
            os.makedirs(history_dir, exist_ok=True)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M")
            archive_path = os.path.join(history_dir, f"local_data_{timestamp}.xlsx")
            shutil.copy2(excel_path, archive_path)
            print(f"🗂️ Архив: {archive_path}")

        os.replace(temp_path, excel_path)

    print("✅ Готово")


if __name__ == "__main__":
    main()
