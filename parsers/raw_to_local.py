from __future__ import annotations

import re
from typing import Iterable

import pandas as pd

from parsers.raw_models import RawAttendance

OUTPUT_COLUMNS = ["תאריך", "כניסה", "יציאה", "אתר", "הערות"]
DATE_RE = re.compile(r"(\d{1,2}[./-]\d{1,2}[./-]\d{2,4})")
TIME_RE = re.compile(r"(?:[01]?\d|2[0-3]):[0-5]\d")


def _normalize_date(value: str) -> str:
    match = DATE_RE.search(value or "")
    if not match:
        return ""
    raw = match.group(1)
    try:
        dt = pd.to_datetime(raw, dayfirst=True)
        return dt.strftime("%d.%m.%Y")
    except Exception:
        return raw


def _to_minutes(value: str) -> int:
    if not value or ":" not in value:
        return 99 * 60
    parts = value.split(":", 1)
    try:
        return int(parts[0]) * 60 + int(parts[1])
    except Exception:
        return 99 * 60




def _warn(meta: dict, message: str) -> None:
    warnings = meta.setdefault("warnings", [])
    if message not in warnings:
        warnings.append(message)


def _get_value(row: dict, key: str) -> str:
    if key in row:
        return str(row.get(key, "") or "").strip()
    reversed_key = key[::-1]
    if reversed_key in row:
        return str(row.get(reversed_key, "") or "").strip()
    english_map = {
        "תאריך": "date",
        "כניסה": "time_in",
        "יציאה": "time_out",
        "אתר": "site",
        "הערות": "notes",
    }
    alt_key = english_map.get(key)
    if alt_key and alt_key in row:
        return str(row.get(alt_key, "") or "").strip()
    return ""


def _canonical_header(header: str) -> str:
    if not header:
        return ""
    if header in ("תאריך", "כניסה", "יציאה", "אתר", "הערות"):
        return header
    reversed_header = header[::-1]
    if reversed_header in ("תאריך", "כניסה", "יציאה", "אתר", "הערות"):
        return reversed_header
    if "סהכ" in header or "סהכ" in reversed_header or 'כ"הס' in header or 'כ"הס' in reversed_header:
        return "סהכ"
    if "ת.כניסה" in header or "ת.כניסה" in reversed_header:
        return "ת.כניסה"
    if "ת.יציאה" in header or "ת.יציאה" in reversed_header:
        return "ת.יציאה"
    return ""


def _row_has_data(date_val: str, time_in: str, time_out: str, site: str, notes: str) -> bool:
    return any([date_val, time_in, time_out, site, notes])


def raw_to_local_df(raw: RawAttendance) -> pd.DataFrame:
    meta = raw.meta or {}
    month = meta.get("month")
    year = meta.get("year")

    records = []
    for row in raw.rows:
        date_raw = _get_value(row, "תאריך")
        time_in = _get_value(row, "כניסה")
        time_out = _get_value(row, "יציאה")
        site = _get_value(row, "אתר")
        notes = _get_value(row, "הערות")

        date_value = _normalize_date(date_raw)
        if not date_value:
            if re.fullmatch(r"\d{1,2}", date_raw or "") and month and year:
                day = int(date_raw)
                date_value = f"{day:02d}.{int(month):02d}.{int(year)}"
            elif date_raw:
                _warn(meta, f"unparsed date: {date_raw}")
            else:
                _warn(meta, "empty date")

        if not _row_has_data(date_value, time_in, time_out, site, notes):
            continue

        if not date_value:
            _warn(meta, "missing date in row, skipped")
            continue

        records.append(
            {
                "תאריך": date_value,
                "כניסה": time_in,
                "יציאה": time_out,
                "אתר": site,
                "הערות": notes,
            }
        )

    if not records:
        raise RuntimeError("No valid rows to write after normalization.")

    df = pd.DataFrame.from_records(records, columns=OUTPUT_COLUMNS)
    df = df.sort_values(by=["תאריך", "כניסה"], key=lambda col: col.map(_to_minutes) if col.name == "כניסה" else col)
    df = df.reset_index(drop=True)
    return df
