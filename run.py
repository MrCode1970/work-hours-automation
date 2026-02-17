import argparse
import os
import shutil
from datetime import datetime

import pandas as pd

from config import load_config
from parsers.normalize_source import normalize_source, save_downloaded_file
from sheets_client import open_spreadsheet, month_sheet_name, get_worksheet
from sync_logic import build_changes_sheet
from ylm_portal import download_excel, open_portal_mobile


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


def _infer_month_from_excel(excel_path: str) -> datetime | None:
    try:
        df = pd.read_excel(excel_path)
    except Exception:
        return None
    if "תאריך" not in df.columns:
        return None
    dates = pd.to_datetime(df["תאריך"], dayfirst=True, errors="coerce")
    dates = dates.dropna()
    if dates.empty:
        return None
    dt = dates.iloc[0]
    return datetime(year=dt.year, month=dt.month, day=1)


def _download_basename() -> str:
    return datetime.now().strftime("%d.%m.%y")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--month", help="Аудит за месяц в формате M.YY (например 12.25)")
    args = parser.parse_args()

    cfg = load_config()

    target_month = _parse_month_arg(args.month) if args.month else None
    sheet_name = _month_sheet_label(target_month) if target_month else month_sheet_name()
    first_day = _first_day_str(target_month) if target_month else None

    def _maybe_infer_month(excel_path: str) -> None:
        nonlocal target_month, sheet_name, first_day
        if target_month is not None:
            return
        inferred = _infer_month_from_excel(excel_path)
        if inferred is not None:
            target_month = inferred
            sheet_name = _month_sheet_label(target_month)
            first_day = _first_day_str(target_month)
        else:
            print("⚠️ Не удалось определить месяц из данных, использую текущий.")

    def _finalize_download(temp_path: str, ext: str) -> str:
        excel_path = normalize_source(temp_path, out_path=cfg["EXCEL_PATH"])
        _maybe_infer_month(excel_path)
        downloads_path = save_downloaded_file(
            temp_path,
            sheet_name,
            ext,
            base_name=_download_basename(),
        )
        raw_path = f"{temp_path}.raw.json"
        if os.path.exists(raw_path):
            target_raw = f"{downloads_path}.raw.json"
            try:
                os.replace(raw_path, target_raw)
            except OSError as exc:
                if getattr(exc, "errno", None) == 18:
                    shutil.move(raw_path, target_raw)
                else:
                    pass
        return excel_path

    # 1. Получаем и нормализуем источник
    source_file = (cfg.get("SOURCE_FILE") or "").strip()
    if source_file:
        if not os.path.exists(source_file):
            raise RuntimeError(f"SOURCE_FILE не найден: {source_file}")
        print(f"▶ SOURCE_FILE — нормализуем: {source_file}")
        excel_path = normalize_source(source_file, out_path=cfg["EXCEL_PATH"])
        _maybe_infer_month(excel_path)
    elif cfg.get("MOBILE_UI"):
        temp_path, ext = open_portal_mobile(
            site_username=cfg["SITE_USERNAME"],
            site_password=cfg["SITE_PASSWORD"],
            headless=cfg["HEADLESS"],
            device_name=cfg["MOBILE_DEVICE"],
            download_timeout_ms=cfg["MANUAL_DOWNLOAD_TIMEOUT_MS"],
        )
        excel_path = _finalize_download(temp_path, ext)
    elif cfg.get("SKIP_DOWNLOAD"):
        excel_path = cfg["EXCEL_PATH"]
        if not os.path.exists(excel_path):
            raise RuntimeError(
                f"SKIP_DOWNLOAD=1 — локальный файл не найден: {excel_path}"
            )
        print(f"⏭️ SKIP_DOWNLOAD=1 — используем локальный Excel: {excel_path}")
        _maybe_infer_month(excel_path)
    else:
        temp_path, ext = download_excel(
            site_username=cfg["SITE_USERNAME"],
            site_password=cfg["SITE_PASSWORD"],
            headless=cfg["HEADLESS"],
            first_day=first_day,
            manual_portal=cfg["MANUAL_PORTAL"],
            manual_download_timeout_ms=cfg["MANUAL_DOWNLOAD_TIMEOUT_MS"],
        )
        excel_path = _finalize_download(temp_path, ext)

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
