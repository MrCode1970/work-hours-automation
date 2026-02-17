from __future__ import annotations

import json
import os
import shutil
from datetime import datetime

from parsers.parse_excel_ylm import parse_excel_ylm
from parsers.parse_pdf_ylm import parse_pdf_ylm
from parsers.raw_to_local import raw_to_local_df
from parsers.write_local_xlsx import write_local_xlsx


def save_downloaded_file(
    source_tmp_path: str,
    sheet_name: str,
    ext: str,
    base_name: str | None = None,
) -> str:
    if not os.path.exists(source_tmp_path):
        raise RuntimeError(f"Скачанный файл не найден: {source_tmp_path}")
    clean_ext = (ext or "").lstrip(".").lower() or "xlsx"
    if base_name:
        base = base_name
    else:
        base = datetime.now().strftime("%d.%m.%y")
    target_dir = os.path.join("downloads", sheet_name)
    os.makedirs(target_dir, exist_ok=True)
    target_path = os.path.join(target_dir, f"{base}.{clean_ext}")
    if os.path.exists(target_path):
        idx = 2
        while True:
            candidate = os.path.join(target_dir, f"{base}_{idx}.{clean_ext}")
            if not os.path.exists(candidate):
                target_path = candidate
                break
            idx += 1
    try:
        os.replace(source_tmp_path, target_path)
    except OSError as exc:
        if getattr(exc, "errno", None) == 18:
            shutil.move(source_tmp_path, target_path)
        else:
            raise
    return target_path


def normalize_source(source_path: str, out_path: str = "local_data.xlsx") -> str:
    ext = os.path.splitext(source_path)[1].lower().lstrip(".")
    if ext in ("xlsx", "xls"):
        raw = parse_excel_ylm(source_path)
    elif ext == "pdf":
        raw = parse_pdf_ylm(source_path)
    else:
        raise RuntimeError(f"Неподдерживаемое расширение: {ext}")

    debug_save_raw = os.getenv("DEBUG_SAVE_RAW", "").strip().lower() in (
        "1",
        "true",
        "yes",
        "y",
        "on",
    )
    if debug_save_raw:
        raw_path = f"{source_path}.raw.json"
        with open(raw_path, "w", encoding="utf-8") as fh:
            json.dump(raw.as_dict(), fh, ensure_ascii=False, indent=2)
        print(f"DEBUG_SAVE_RAW: saved {raw_path}")

    df = raw_to_local_df(raw)
    write_local_xlsx(df, out_path)
    return out_path
