import argparse
import os
import shutil
from datetime import datetime

from config import load_config
from sheets_client import open_spreadsheet, month_sheet_name, get_worksheet
from sync_logic import build_changes_sheet
from ylm_portal import download_excel


def _parse_month_arg(raw: str) -> datetime:
    """
    Ожидается формат M.YY (например 12.25).
    """
    parts = raw.strip().split(".")
    if len(parts) != 2:
        raise ValueError("Ожидается формат M.YY (например 12.25)")
    month = int(parts[0])
    year = 2000 + int(parts[1])
    return datetime(year=year, month=month, day=1)


def _month_sheet_label(dt: datetime) -> str:
    return f"{dt.month}.{dt.strftime('%y')}"


def _first_day_str(dt: datetime) -> str:
    return f"01/{dt.strftime('%m/%Y')}"


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--month", help="Аудит за месяц в формате M.YY (например 12.25)")
    args = parser.parse_args()

    cfg = load_config()

    target_month = _parse_month_arg(args.month) if args.month else None
    sheet_name = _month_sheet_label(target_month) if target_month else month_sheet_name()
    first_day = _first_day_str(target_month) if target_month else None

    # 1. Получаем Excel
    history_dir = "history"
    os.makedirs(history_dir, exist_ok=True)

    if target_month:
        excel_path = os.path.join(history_dir, f"{sheet_name}.xlsx")
        if os.path.exists(excel_path):
            print(f"📦 Используем архив: {excel_path}")
        elif cfg.get("SKIP_DOWNLOAD"):
            raise RuntimeError(f"Архив за {sheet_name} не найден: {excel_path}")
        else:
            temp_path = f"{excel_path}.new"
            download_excel(
                site_username=cfg["SITE_USERNAME"],
                site_password=cfg["SITE_PASSWORD"],
                excel_path=temp_path,
                headless=cfg["HEADLESS"],
                first_day=first_day,
                manual_portal=cfg["MANUAL_PORTAL"],
                manual_download_timeout_ms=cfg["MANUAL_DOWNLOAD_TIMEOUT_MS"],
            )
            os.replace(temp_path, excel_path)
            print(f"📦 Архив сохранён: {excel_path}")
    else:
        if cfg.get("SKIP_DOWNLOAD"):
            excel_path = cfg["EXCEL_PATH"]
            print(f"⏭️ SKIP_DOWNLOAD=1 — используем локальный Excel: {excel_path}")
        else:
            excel_path = cfg["EXCEL_PATH"]
            temp_path = f"{excel_path}.new"

            prev_month = datetime.now().replace(day=1)
            prev_month = prev_month.replace(month=12, year=prev_month.year - 1) if prev_month.month == 1 else prev_month.replace(month=prev_month.month - 1)
            prev_label = _month_sheet_label(prev_month)
            prev_archive = os.path.join(history_dir, f"{prev_label}.xlsx")
            if os.path.exists(excel_path) and not os.path.exists(prev_archive):
                shutil.copy2(excel_path, prev_archive)
                print(f"🗂️ Архив за прошлый месяц: {prev_archive}")

            download_excel(
                site_username=cfg["SITE_USERNAME"],
                site_password=cfg["SITE_PASSWORD"],
                excel_path=temp_path,
                headless=cfg["HEADLESS"],
                first_day=first_day,
                manual_portal=cfg["MANUAL_PORTAL"],
                manual_download_timeout_ms=cfg["MANUAL_DOWNLOAD_TIMEOUT_MS"],
            )
            os.replace(temp_path, excel_path)

    # 2. Открываем Google Sheets
    spreadsheet = open_spreadsheet(
        gsheet_id=cfg["GSHEET_ID"],
        google_json_file=cfg["GOOGLE_JSON_FILE"],
    )

    # 3. Получаем основной лист месяца (эталон)
    try:
        worksheet = get_worksheet(spreadsheet, sheet_name)
    except Exception as exc:
        raise RuntimeError(f"Лист {sheet_name} не найден.") from exc

    # 4. Строим лист "Изменения M.YY"
    changes_found = build_changes_sheet(
        spreadsheet=spreadsheet,
        base_ws=worksheet,
        sheet_name=sheet_name,
        excel_path=excel_path,
    )

    print("✅ Готово")


if __name__ == "__main__":
    main()
