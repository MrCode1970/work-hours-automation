from __future__ import annotations

import os
import re

import pandas as pd

from parsers.raw_models import RawAttendance

REQUIRED_COLUMNS = ["תאריך", "כניסה", "יציאה"]


def _stringify(value) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    if isinstance(value, pd.Timestamp):
        return value.strftime("%d.%m.%Y")
    return str(value).strip()


def _extract_month_year_from_path(path: str) -> tuple[int | None, int | None]:
    match = re.search(r"(?<!\d)(0?[1-9]|1[0-2])[./_-](\d{2,4})(?!\d)", path or "")
    if not match:
        return None, None
    month = int(match.group(1))
    raw_year = match.group(2)
    year = int(raw_year) if len(raw_year) == 4 else 2000 + int(raw_year)
    return month, year


def parse_excel_ylm(source_path: str) -> RawAttendance:
    df = pd.read_excel(source_path)
    if not all(c in df.columns for c in REQUIRED_COLUMNS):
        raise RuntimeError("Excel не содержит ожидаемые колонки: תאריך, כניסה, יציאה")

    headers = [str(c).strip() for c in df.columns]
    rows: list[dict[str, str]] = []
    for _, row in df.iterrows():
        row_dict: dict[str, str] = {}
        for col in df.columns:
            row_dict[str(col).strip()] = _stringify(row[col])
        rows.append(row_dict)

    month, year = _extract_month_year_from_path(os.path.basename(source_path))
    raw = RawAttendance(
        meta={
            "source_path": source_path,
            "source_kind": "xlsx",
            "parser_mode": "XLSX_TABLE",
            "year": year,
            "month": month,
            "month_he": None,
            "headers": headers,
            "pages": None,
            "generated_at": None,
            "warnings": [],
        },
        rows=rows,
    )
    return raw
