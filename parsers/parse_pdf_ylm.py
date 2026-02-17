from __future__ import annotations

import os
import re
from dataclasses import dataclass
from typing import Iterable, Iterator

import pandas as pd
import pdfplumber

from parsers.raw_models import RawAttendance

DATE_RE = re.compile(r"(\d{1,2}[./-]\d{1,2}[./-]\d{2,4})")
TIME_RE = re.compile(r"(?:[01]?\d|2[0-3]):[0-5]\d")
DAY_RE = re.compile(r"\b([0-2]?\d|3[01])\b")
_BIDI_CHARS_RE = re.compile(r"[\u200e\u200f\u202a-\u202e\u2066-\u2069]")
HEBREW_RE = re.compile(r"[\u0590-\u05ff]")

MONTH_MAP = {
    "ינואר": 1,
    "פברואר": 2,
    "מרץ": 3,
    "מארס": 3,
    "אפריל": 4,
    "מאי": 5,
    "יוני": 6,
    "יולי": 7,
    "אוגוסט": 8,
    "ספטמבר": 9,
    "אוקטובר": 10,
    "נובמבר": 11,
    "דצמבר": 12,
}

IGNORED_TIMES = {"00:00", "00:01"}

MONTH_MAP_REVERSED = {name[::-1]: (number, name) for name, number in MONTH_MAP.items()}

SERVICE_LINE_PATTERNS = (
    "סה\"כ",
    "סה״כ",
    "סהכ",
    "סיכום",
    "חתימה",
    "דו\"ח",
    "דוח",
    "שם",
    "עובד",
    "טווח",
    "חודש",
    "שנה",
    "תאריך",
)

CANONICAL_ROW_KEYS = ("date", "time_in", "time_out", "site", "notes")
CANONICAL_META_KEYS = (
    "report_date",
    "month_total_hours",
    "month_total_days",
    "trips",
    "total_row",
)


@dataclass
class _ParseIssue:
    line: str
    reason: str


def _normalize_cell(value) -> str:
    return str(value or "").strip()


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


def _clean_text(text: str) -> str:
    text = text or ""
    text = _BIDI_CHARS_RE.sub("", text)
    return text


def _normalize_space(text: str) -> str:
    return re.sub(r"\s+", " ", _clean_text(text)).strip()


def _get_bool_env(name: str, default: bool) -> bool:
    raw = os.getenv(name, "1" if default else "0").strip().lower()
    return raw in ("1", "true", "yes", "y", "on")


def _normalize_display_text(text: str) -> str:
    text = text or ""
    if _get_bool_env("PDF_STRIP_BIDI", True):
        text = _BIDI_CHARS_RE.sub("", text)
    text = re.sub(r"\s+", " ", text).strip()
    if _get_bool_env("PDF_REVERSE_HEBREW", False):
        if HEBREW_RE.search(text) and not re.search(r"[A-Za-z]", text):
            text = text[::-1]
    return text


def _normalize_time(value: str) -> str:
    raw = _normalize_space(value).replace(" ", "")
    if not raw:
        return ""
    match = re.search(r"(\d{1,2})[:.](\d{2})", raw)
    if not match:
        return ""
    hours = int(match.group(1))
    minutes = int(match.group(2))
    if hours < 0 or hours > 23 or minutes < 0 or minutes > 59:
        return ""
    return f"{hours:02d}:{minutes:02d}"


def _extract_month_year(text: str) -> tuple[int | None, int | None, str | None]:
    text = _clean_text(text)
    lowered = text.lower()
    month = None
    month_he = None
    for name, number in MONTH_MAP.items():
        if name in lowered:
            month = number
            month_he = name
            break
    if month is None:
        for reversed_name, payload in MONTH_MAP_REVERSED.items():
            if reversed_name in lowered:
                number, original = payload
                month = number
                month_he = original
                break

    year = None
    year_candidates: list[int] = []
    for pattern in (
        r"(?:שנה|הנש)\s*[:\-]?\s*(\d{2,4})",
        r"(\d{2,4})\s*[:\-]?\s*(?:שנה|הנש)",
    ):
        for match in re.finditer(pattern, lowered):
            raw = match.group(1)
            value = int(raw) if len(raw) == 4 else 2000 + int(raw)
            if 2000 <= value <= 2100:
                year_candidates.append(value)
    if year_candidates:
        year = year_candidates[0]
    else:
        year_match = re.search(r"\b(20\d{2})\b", lowered)
        if year_match:
            year = int(year_match.group(1))

    if month is None:
        month_match = re.search(r"\b(1[0-2]|0?[1-9])[./](\d{2})\b", lowered)
        if month_match and year is None:
            month = int(month_match.group(1))
            year = 2000 + int(month_match.group(2))
        elif month_match:
            month = int(month_match.group(1))

    return month, year, month_he


def _extract_day(line: str) -> int | None:
    line_wo_times = TIME_RE.sub(" ", line)
    match = DAY_RE.search(line_wo_times)
    if not match:
        return None
    day = int(match.group(1))
    if day < 1 or day > 31:
        return None
    return day


def _times_in_text(text: str) -> list[tuple[str, int, int]]:
    return [(m.group(0), m.start(), m.end()) for m in TIME_RE.finditer(text or "")]


def _select_time_pair(times: list[str]) -> tuple[str, str] | None:
    clean_times = [t for t in times if t not in IGNORED_TIMES]
    if len(clean_times) < 2:
        return None
    if len(clean_times) == 2:
        def _to_minutes(value: str) -> int:
            if ":" not in value:
                return -1
            hours, minutes = value.split(":", 1)
            return int(hours) * 60 + int(minutes)

        first, second = clean_times[0], clean_times[1]
        if _to_minutes(first) <= _to_minutes(second):
            return first, second
        return second, first

    parsed = []
    for t in clean_times:
        if ":" not in t:
            continue
        hours, minutes = t.split(":", 1)
        parsed.append((t, int(hours) * 60 + int(minutes)))

    parsed.sort(key=lambda item: item[1])

    candidates: list[tuple[int, str, str]] = []
    for i in range(len(parsed)):
        for j in range(i + 1, len(parsed)):
            t_in, m_in = parsed[i]
            t_out, m_out = parsed[j]
            if m_out <= m_in:
                continue
            duration = m_out - m_in
            if 3 * 60 <= duration <= 16 * 60:
                score = abs(duration - 8 * 60)
                candidates.append((score, t_in, t_out))
    if not candidates:
        return None
    candidates.sort(key=lambda item: item[0])
    _, t_in, t_out = candidates[0]
    return t_in, t_out


def _is_service_line(line: str) -> bool:
    lowered = (line or "").lower()
    if not lowered.strip():
        return True
    return any(token in lowered for token in SERVICE_LINE_PATTERNS)


def _record(date: str, time_in: str, time_out: str, site: str = "", notes: str = "") -> dict:
    return {"תאריך": date, "כניסה": time_in, "יציאה": time_out, "אתר": site, "הערות": notes}


def _parse_type_b_text(pages: Iterable, month: int | None, year: int | None) -> tuple[list[dict], list[_ParseIssue]]:
    records: list[dict] = []
    bad: list[_ParseIssue] = []
    for page in pages:
        text = page.extract_text() or ""
        for raw_line in text.splitlines():
            line = _clean_text(raw_line)
            line = line.replace("\xad", " ")
            line = re.sub(r"\s+", " ", line).strip()
            if _is_service_line(line):
                continue

            times_with_pos = _times_in_text(line)
            times = [t for t, _, _ in times_with_pos]
            if len(times) < 2:
                bad.append(_ParseIssue(line=line, reason="нет 2 времен"))
                continue

            date_match = DATE_RE.search(line)
            date_value = _normalize_date(line)
            day = None
            if not date_value:
                day = _extract_day(line)
                if day is None or month is None or year is None:
                    bad.append(_ParseIssue(line=line, reason="нет даты/контекста месяца"))
                    continue
                date_value = f"{day:02d}"

            if len(times) >= 3 and any(token in line for token in ("סה\"כ", "סה״כ", "סהכ")):
                times = times[:2]

            pair = _select_time_pair(times)
            if not pair:
                bad.append(_ParseIssue(line=line, reason="нет корректной пары времен"))
                continue
            time_in, time_out = pair

            site = ""
            notes = ""
            is_table_like = date_match is not None and len(times) >= 2
            if not is_table_like:
                if date_match:
                    site = line[: date_match.start()].strip(" -|")
                elif day is not None:
                    day_match = DAY_RE.search(TIME_RE.sub(" ", line))
                    if day_match:
                        site = line[: day_match.start()].strip(" -|")

                if times_with_pos:
                    _, _, last_end = times_with_pos[min(1, len(times_with_pos) - 1)]
                    tail = line[last_end:].strip()
                    if tail and not any(token in tail for token in ("סה\"כ", "סה״כ", "סהכ")):
                        notes = tail

            records.append(_record(date_value, time_in, time_out, site, notes))

    return records, bad


def _iter_table_rows(tables: Iterable[list[list[str]]]) -> Iterator[list[str]]:
    for table in tables:
        if not table:
            continue
        for row in table:
            if row is None:
                continue
            yield [_normalize_cell(cell) for cell in row]


def _parse_type_a_tables(
    pages: Iterable,
    *,
    month: int | None,
    year: int | None,
) -> tuple[list[dict[str, str]], str, list[_ParseIssue], list[str]]:
    return _parse_type_a_tables_canonical(pages, month=month, year=year)


def _header_key(cell: str) -> str:
    cleaned = _normalize_space(cell).replace('"', "").replace("״", "").replace("'", "")
    if not cleaned:
        return ""
    lowered = cleaned.lower()
    reversed_cleaned = cleaned[::-1].lower()

    def _has(token: str) -> bool:
        return token in lowered or token in reversed_cleaned

    if _has("תאריך"):
        return "date"
    if _has("יום"):
        return "day"
    if _has("כניסה") or _has("ת.כניסה"):
        return "time_in"
    if _has("יציאה") or _has("ת.יציאה"):
        return "time_out"
    if _has("אתר"):
        return "site"
    if _has("הערות"):
        return "notes"
    if _has("סהכ") or _has("סה\"כ") or _has("סה״כ"):
        return "total"
    return ""


def _extract_day_from_cell(value: str) -> int | None:
    cleaned = _normalize_space(value)
    if not cleaned:
        return None
    match = DAY_RE.search(cleaned)
    if not match:
        return None
    day = int(match.group(1))
    if day < 1 or day > 31:
        return None
    return day


def _build_date(value: str, month: int | None, year: int | None) -> str:
    full = _normalize_date(value)
    if full:
        return full
    day = _extract_day_from_cell(value)
    if day is None or month is None or year is None:
        return ""
    return f"{day:02d}.{int(month):02d}.{int(year)}"


def _canonical_row(
    date_value: str,
    time_in: str,
    time_out: str,
    site: str,
    notes: str,
) -> dict[str, str]:
    return {
        "date": date_value,
        "time_in": time_in,
        "time_out": time_out,
        "site": site,
        "notes": notes,
    }


def _parse_type_a_tables_canonical(
    pages: Iterable,
    *,
    month: int | None,
    year: int | None,
) -> tuple[list[dict[str, str]], str, list[_ParseIssue], list[str]]:
    records: list[dict[str, str]] = []
    bad: list[_ParseIssue] = []
    headers: list[str] = []
    total_row_value = ""

    for page in pages:
        tables = page.extract_tables() or []
        for table in tables:
            normalized_rows = []
            for row in table or []:
                if row is None:
                    continue
                normalized_rows.append([_normalize_cell(cell) for cell in row])
            if len(normalized_rows) < 2:
                continue

            header_idx = None
            header_hits = 0
            for idx, row in enumerate(normalized_rows):
                row_text = " ".join(_normalize_space(cell) for cell in row if cell)
                hits = 0
                for token in ("תאריך", "יום", "כניסה", "יציאה", "אתר", "הערות", "סהכ"):
                    if token in row_text or token[::-1] in row_text:
                        hits += 1
                if hits > header_hits:
                    header_hits = hits
                    header_idx = idx

            if header_idx is None or header_hits < 2:
                bad.append(_ParseIssue(line="header_not_found", reason="no header row"))
                continue

            header_row = normalized_rows[header_idx]
            table_headers = [cell for cell in header_row]
            if not any(table_headers):
                table_headers = [f"COL_{idx+1}" for idx in range(len(header_row))]
            else:
                table_headers = [
                    cell if cell else f"COL_{idx+1}" for idx, cell in enumerate(table_headers)
                ]
            data_rows = normalized_rows[header_idx + 1 :]

            if not headers:
                headers = list(table_headers)

            key_map = [_header_key(cell) for cell in table_headers]
            key_indices: dict[str, int] = {}
            for idx, key in enumerate(key_map):
                if key and key not in key_indices:
                    key_indices[key] = idx

            for row in data_rows:
                if not any(cell for cell in row):
                    continue

                date_cell = row[key_indices["date"]] if "date" in key_indices and key_indices["date"] < len(row) else ""
                time_in_cell = row[key_indices["time_in"]] if "time_in" in key_indices and key_indices["time_in"] < len(row) else ""
                time_out_cell = row[key_indices["time_out"]] if "time_out" in key_indices and key_indices["time_out"] < len(row) else ""
                site_cell = row[key_indices["site"]] if "site" in key_indices and key_indices["site"] < len(row) else ""
                notes_cell = row[key_indices["notes"]] if "notes" in key_indices and key_indices["notes"] < len(row) else ""

                total_cell = ""
                if "total" in key_indices and key_indices["total"] < len(row):
                    total_cell = _normalize_space(row[key_indices["total"]])

                date_value = _build_date(date_cell, month, year)
                time_in = _normalize_time(time_in_cell)
                time_out = _normalize_time(time_out_cell)
                site = _normalize_display_text(site_cell)
                notes = _normalize_display_text(notes_cell)

                total_col_idx = key_indices.get("total")
                time_candidates = []
                for idx, cell in enumerate(row):
                    if total_col_idx is not None and idx == total_col_idx:
                        continue
                    norm_time = _normalize_time(cell)
                    if norm_time:
                        time_candidates.append((idx, norm_time))

                if not time_in and time_candidates:
                    time_candidates_sorted = sorted(time_candidates, key=lambda item: item[1])
                    time_in = time_candidates_sorted[0][1]

                if not time_out and time_candidates:
                    other_times = [t for _, t in time_candidates if t != time_in]
                    if other_times:
                        time_out = max(other_times)

                has_regular_data = any([time_in, time_out, site, notes])

                if total_cell and not has_regular_data:
                    if not total_row_value:
                        total_row_value = total_cell
                    continue

                if not has_regular_data:
                    continue

                records.append(_canonical_row(date_value, time_in, time_out, site, notes))

    return records, total_row_value, bad, headers


def _extract_meta_from_text(all_text: str) -> dict[str, str]:
    meta = {key: "" for key in CANONICAL_META_KEYS}
    text = _clean_text(all_text)

    def _search(patterns: list[str]) -> re.Match | None:
        for pattern in patterns:
            match = re.search(pattern, text)
            if match:
                return match
        return None

    report_match = _search(
        [
            r"(?:תאריך(?:\s*ה?דוח)?|ה?דוח\s*תאריך)\s*[:\-]?\s*(\d{1,2}[./-]\d{1,2}[./-]\d{2,4})",
            r"(\d{1,2}[./-]\d{1,2}[./-]\d{2,4})\s*[:\-]?\s*(?:חוד\s*ךיראת|ךיראת\s*חוד)",
        ]
    )
    if report_match:
        meta["report_date"] = _normalize_date(report_match.group(1))

    hours_match = _search(
        [
            r"שעות\s*עבודה\s*בפועל\s*[:\-]?\s*([0-9]{1,3}:[0-5]\d)",
            r"([0-9]{1,3}:[0-5]\d)\s*[:\-]?\s*לעופב\s*הדובע\s*תועש",
        ]
    )
    if hours_match:
        meta["month_total_hours"] = hours_match.group(1)

    days_match = _search(
        [
            r"ימי\s*עבודה\s*בפועל\s*[:\-]?\s*([0-9]{1,2})",
            r"([0-9]{1,2})\s*[:\-]?\s*לעופב\s*הדובע\s*ימי",
        ]
    )
    if days_match:
        meta["month_total_days"] = days_match.group(1)

    trips_match = _search(
        [
            r"נסיעות\s*[:\-]?\s*([0-9]+(?:[.,][0-9]+)?)",
            r"([0-9]+(?:[.,][0-9]+)?)\s*[:\-]?\s*תועיסנ",
        ]
    )
    if trips_match:
        meta["trips"] = trips_match.group(1).replace(",", ".")

    return meta




def _debug_print(
    mode: str,
    pages_count: int,
    records: list[dict],
    bad: list[_ParseIssue],
    month: int | None,
    year: int | None,
    headers: list[str] | None,
) -> None:
    print(f"DEBUG_PDF: mode={mode}")
    print(f"DEBUG_PDF: pages={pages_count}")
    print(f"DEBUG_PDF: year={year} month={month}")
    if headers is not None:
        print(f"DEBUG_PDF: headers={headers[:10]}")
    print(f"DEBUG_PDF: records={len(records)}")
    print("DEBUG_PDF: sample:")
    for rec in records[:10]:
        if "תאריך" in rec:
            print(f"  {rec.get('תאריך')} | {rec.get('כניסה')} | {rec.get('יציאה')} | {rec.get('אתר')} | {rec.get('הערות')}")
        elif "date" in rec:
            print(f"  {rec.get('date')} | {rec.get('time_in')} | {rec.get('time_out')} | {rec.get('site')} | {rec.get('notes')}")
        else:
            print(f"  {rec}")
    if bad:
        print("DEBUG_PDF: bad lines:")
        for issue in bad[:10]:
            print(f"  {issue.reason}: {issue.line}")


def parse_pdf_ylm(source_path: str) -> RawAttendance:
    debug = os.getenv("DEBUG_PDF", "").strip().lower() in ("1", "true", "yes", "y", "on")

    try:
        with pdfplumber.open(source_path) as pdf:
            pages = list(pdf.pages)
            all_text = "\n".join((page.extract_text() or "") for page in pages)
            month, year, month_he = _extract_month_year(all_text)

            warnings = []
            if month is None or year is None:
                warnings.append("month/year not found in PDF text")

            table_records, total_row_value, table_bad, table_headers = _parse_type_a_tables(
                pages, month=month, year=year
            )
            if table_records:
                meta = _extract_meta_from_text(all_text)
                if total_row_value:
                    meta["total_row"] = total_row_value
                else:
                    meta["total_row"] = meta.get("total_row", "")
                if debug:
                    _debug_print("TYPE_A_TABLE", len(pages), table_records, table_bad, month, year, table_headers)
                return RawAttendance(
                    meta=meta,
                    rows=table_records,
                )

            text_records, text_bad = _parse_type_b_text(pages, month, year)
            if not text_records:
                reason = "Не удалось извлечь строки из PDF."
                if month is None or year is None:
                    reason += " Не найден месяц/год в тексте PDF."
                raise RuntimeError(reason)
            headers = ["תאריך", "כניסה", "יציאה", "אתר", "הערות"]
            if debug:
                _debug_print("TYPE_B_TEXT", len(pages), text_records, text_bad, month, year, headers)
            return RawAttendance(
                meta={
                    "source_path": source_path,
                    "source_kind": "pdf",
                    "parser_mode": "TYPE_B_TEXT",
                    "year": year,
                    "month": month,
                    "month_he": month_he,
                    "headers": headers,
                    "pages": len(pages),
                    "generated_at": None,
                    "warnings": warnings,
                },
                rows=text_records,
            )
    except RuntimeError:
        raise
    except Exception as exc:
        raise RuntimeError(f"Не удалось прочитать PDF: {exc}") from exc
