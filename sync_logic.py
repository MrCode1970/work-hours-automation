from __future__ import annotations

from datetime import datetime
import pandas as pd


def _normalize_time(value, *, empty_as_zero: bool = False) -> str:
    """
    Нормализует время к HH:MM.
    Принимает строки вида "7:00", "07:00", "07:00:00", а также NaN/None.
    """
    if value is None:
        return "00:00" if empty_as_zero else ""
    s = str(value).strip()
    if s == "" or s.lower() == "nan":
        return "00:00" if empty_as_zero else ""
    s = s[:8]
    if ":" not in s:
        return "00:00" if empty_as_zero else ""

    parts = s.split(":")
    if len(parts) < 2:
        return ""
    h, m = parts[0], parts[1]
    try:
        hh = int(h)
        mm = int(m)
        if not (0 <= hh <= 23 and 0 <= mm <= 59):
            return ""
        return f"{hh:02d}:{mm:02d}"
    except Exception:
        return "00:00" if empty_as_zero else ""


def _time_to_minutes(t: str) -> int:
    if not t or ":" not in t:
        return 0
    h, m = t.split(":")[:2]
    return int(h) * 60 + int(m)


def _format_time_for_sheet(v, *, empty_as_zero: bool = False) -> str:
    """
    Нормализует время для записи в Google Sheets без апострофов.
    """
    return _normalize_time(v, empty_as_zero=empty_as_zero)


def _color_red():
    return {"red": 0.80, "green": 0.00, "blue": 0.00}


def _color_green():
    return {"red": 0.00, "green": 0.60, "blue": 0.00}


def _bg_my():
    # мягкий серый
    return {"red": 0.93, "green": 0.93, "blue": 0.93}


def _bg_site():
    # мягкий жёлтый
    return {"red": 1.00, "green": 0.98, "blue": 0.85}


def _delete_worksheet_if_exists(spreadsheet, title: str) -> None:
    try:
        ws = spreadsheet.worksheet(title)
        spreadsheet.del_worksheet(ws)
    except Exception:
        return


def build_changes_sheet(spreadsheet, base_ws, sheet_name: str, excel_path: str) -> bool:
    """
    Создаёт/пересоздаёт лист "Изменения M.YY" (состояние расхождений).
    Если расхождений нет — лист удаляется (или не создаётся).

    Столбцы листа изменений строго:
    Дата | Факт (Вход/Выход) | Табель (Вход/Выход) | Разница
    (бонусы идут отдельной строкой без даты)
    """

    changes_title = f"Изменения {sheet_name}"

    # 1) Считаем Excel (сайт)
    df = pd.read_excel(excel_path)
    if not all(c in df.columns for c in ["תאריך", "כניסה", "יציאה"]):
        raise RuntimeError("Excel не содержит ожидаемые колонки: תאריך, כניסה, יציאה")
    df = df[["תאריך", "כניסה", "יציאה"]].dropna(subset=["תאריך"])

    # date_obj -> [(site_in, site_out), ...]
    site_by_date: dict[datetime, list[tuple[str, str]]] = {}
    for _, r in df.iterrows():
        raw_date = str(r["תאריך"]).split()[0]
        try:
            d = pd.to_datetime(raw_date, dayfirst=True)
        except Exception:
            continue

        site_in = _format_time_for_sheet(r["כניסה"])
        site_out = _format_time_for_sheet(r["יציאה"])
        if site_in == "" and site_out == "":
            continue

        key = d.normalize()
        site_by_date.setdefault(key, []).append((site_in, site_out))

    # 2) Считаем базовую таблицу (твои часы — эталон)
    base_values = base_ws.get_values("B:L")

    # Индекс по дате из base:
    # date_str -> (row_num, my_in_1, my_out_1, my_in_2, my_out_2)
    # В base дата в столбце B (index 1), вход/выход в C/D (2/3), бонус в K/L (10/11)
    base_by_date: dict[str, tuple[int, str, str, str, str]] = {}
    for idx, row in enumerate(base_values):
        if len(row) < 1:
            continue
        date_cell = str(row[0]).strip()
        if not date_cell:
            continue

        my_in_1 = _format_time_for_sheet(row[1] if len(row) > 1 else "")
        my_out_1 = _format_time_for_sheet(row[2] if len(row) > 2 else "")
        my_in_2 = _format_time_for_sheet(row[9] if len(row) > 9 else "")
        my_out_2 = _format_time_for_sheet(row[10] if len(row) > 10 else "")
        row_num = idx + 1  # 1-based for Sheets API

        base_by_date[date_cell] = (row_num, my_in_1, my_out_1, my_in_2, my_out_2)

    # 3) Собираем изменения
    # Каждая строка:
    # [date, my_in, my_out, site_in, site_out, diff_formula, cmp_in, cmp_out]
    changes_rows = []
    base_updates = []
    updated_my_cache: dict[str, tuple[str, str, str, str]] = {}

    def _row_for_interval(date_label: str, my_in: str, my_out: str, site_in: str, site_out: str):
        my_in = _format_time_for_sheet(my_in)
        my_out = _format_time_for_sheet(my_out)
        site_in = _format_time_for_sheet(site_in)
        site_out = _format_time_for_sheet(site_out)

        cmp_in = 0
        cmp_out = 0
        if my_in and site_in and my_in != site_in:
            cmp_in = -1 if _time_to_minutes(site_in) < _time_to_minutes(my_in) else 1
        if my_out and site_out and my_out != site_out:
            cmp_out = -1 if _time_to_minutes(site_out) < _time_to_minutes(my_out) else 1

        return [date_label, my_in, my_out, site_in, site_out, "", cmp_in, cmp_out]

    for date_key, intervals in site_by_date.items():
        d = pd.to_datetime(date_key)
        date_variants = [d.strftime("%d.%m.%Y"), d.strftime("%d/%m/%Y")]

        intervals_sorted = sorted(
            intervals,
            key=lambda x: (x[0] == "", _time_to_minutes(x[0])),
        )
        site_in_1, site_out_1 = intervals_sorted[0] if len(intervals_sorted) > 0 else ("", "")
        site_in_2, site_out_2 = intervals_sorted[1] if len(intervals_sorted) > 1 else ("", "")
        if len(intervals_sorted) > 2:
            print(f"⚠️ Дата {date_variants[0]}: найдено интервалов {len(intervals_sorted)}, используем первые два.")

        # Найдём дату в основном листе (по одному из вариантов формата)
        base_date = None
        row_num = None
        my_in_1 = ""
        my_out_1 = ""
        my_in_2 = ""
        my_out_2 = ""
        for dv in date_variants:
            if dv in base_by_date:
                base_date = dv
                row_num, my_in_1, my_out_1, my_in_2, my_out_2 = base_by_date[dv]
                break

        if base_date is None or row_num is None:
            # даты нет в твоём листе — это не "изменение"
            continue

        if base_date in updated_my_cache:
            my_in_1, my_out_1, my_in_2, my_out_2 = updated_my_cache[base_date]

        changed_base = False
        if my_in_1 == "" and site_in_1 != "":
            base_updates.append({"range": f"C{row_num}", "values": [[site_in_1]]})
            my_in_1 = site_in_1
            changed_base = True
        if my_out_1 == "" and site_out_1 != "":
            base_updates.append({"range": f"D{row_num}", "values": [[site_out_1]]})
            my_out_1 = site_out_1
            changed_base = True
        if my_in_2 == "" and site_in_2 != "":
            base_updates.append({"range": f"K{row_num}", "values": [[site_in_2]]})
            my_in_2 = site_in_2
            changed_base = True
        if my_out_2 == "" and site_out_2 != "":
            base_updates.append({"range": f"L{row_num}", "values": [[site_out_2]]})
            my_out_2 = site_out_2
            changed_base = True

        if changed_base:
            updated_my_cache[base_date] = (my_in_1, my_out_1, my_in_2, my_out_2)

        main_diff = (my_in_1 != site_in_1) or (my_out_1 != site_out_1)
        bonus_diff = (my_in_2 != site_in_2) or (my_out_2 != site_out_2)
        has_bonus = bool(my_in_2 or my_out_2 or site_in_2 or site_out_2)

        if not (main_diff or bonus_diff):
            continue

        changes_rows.append(_row_for_interval(base_date, my_in_1, my_out_1, site_in_1, site_out_1))
        if has_bonus:
            changes_rows.append(_row_for_interval("", my_in_2, my_out_2, site_in_2, site_out_2))

    # 4) Применяем дозаполнения в основной таблице одним пакетом
    if base_updates:
        try:
            base_ws.batch_update(base_updates, value_input_option="USER_ENTERED")
        except AttributeError:
            for u in base_updates:
                base_ws.update(u["range"], u["values"], value_input_option="USER_ENTERED")

    # 5) Если расхождений нет — удалить лист и выйти
    if not changes_rows:
        _delete_worksheet_if_exists(spreadsheet, changes_title)
        print(f"✅ Расхождений нет — лист '{changes_title}' удалён/не создан.")
        return False

    # 6) Пересоздать лист изменений
    _delete_worksheet_if_exists(spreadsheet, changes_title)
    ws = spreadsheet.add_worksheet(title=changes_title, rows=len(changes_rows) + 10, cols=6)

    # 6) A1 и заголовки
    ws.update("A1", [[f"Дата изменений: {datetime.now().strftime('%d.%m.%Y')}"]])

    header_rows = [
        ["Дата", "Факт", "", "Табель", "", "Разница"],
        ["", "Вход", "Выход", "Вход", "Выход", ""],
    ]
    ws.update("A3:F4", header_rows, value_input_option="USER_ENTERED")
    spreadsheet.batch_update(
        {
            "requests": [
                {
                    "mergeCells": {
                        "range": {
                            "sheetId": ws.id,
                            "startRowIndex": 2,
                            "endRowIndex": 3,
                            "startColumnIndex": 1,
                            "endColumnIndex": 3,
                        },
                        "mergeType": "MERGE_ALL",
                    }
                },
                {
                    "mergeCells": {
                        "range": {
                            "sheetId": ws.id,
                            "startRowIndex": 2,
                            "endRowIndex": 3,
                            "startColumnIndex": 3,
                            "endColumnIndex": 5,
                        },
                        "mergeType": "MERGE_ALL",
                    }
                },
            ]
        }
    )

    # 7) Данные одним блоком
    start_row = 5
    values_block = []
    for idx, rr in enumerate(changes_rows):
        row_num = start_row + idx
        diff_formula = (
            f'=ЕСЛИ(И(B{row_num}<>"";C{row_num}<>"";D{row_num}<>"";E{row_num}<>"");'
            f'(E{row_num}-D{row_num})-(C{row_num}-B{row_num});"")'
        )
        if rr[0] == "":
            fact_in = _format_time_for_sheet(rr[1], empty_as_zero=True)
            fact_out = _format_time_for_sheet(rr[2], empty_as_zero=True)
            site_in = _format_time_for_sheet(rr[3], empty_as_zero=True)
            site_out = _format_time_for_sheet(rr[4], empty_as_zero=True)
            values_block.append([rr[0], fact_in, fact_out, site_in, site_out, diff_formula])
        else:
            values_block.append([rr[0], rr[1], rr[2], rr[3], rr[4], diff_formula])

    end_row = start_row + len(values_block) - 1
    ws.update(
        f"A{start_row}:F{end_row}",
        values_block,
        value_input_option="USER_ENTERED",
    )

    # 8) Фон групп (как у тебя по образцу)
    ws.batch_format(
        [
            {"range": "A3:F4", "format": {"textFormat": {"bold": True}}},
            {"range": f"B4:C{end_row}", "format": {"backgroundColor": _bg_my()}},
            {"range": f"D4:E{end_row}", "format": {"backgroundColor": _bg_site()}},
            {"range": f"B{start_row}:E{end_row}", "format": {"numberFormat": {"type": "TIME", "pattern": "hh:mm"}}},
            {"range": f"F{start_row}:F{end_row}", "format": {"numberFormat": {"type": "TIME", "pattern": "[h]:mm"}}},
        ]
    )

    # 9) Окраска текста: сайт и разница
    for idx, rr in enumerate(changes_rows):
        row_num = start_row + idx

        cmp_in = rr[6]
        cmp_out = rr[7]
        if cmp_in != 0:
            c = _color_red() if cmp_in < 0 else _color_green()
            ws.format(f"D{row_num}", {"textFormat": {"foregroundColor": c}})
        if cmp_out != 0:
            c = _color_red() if cmp_out < 0 else _color_green()
            ws.format(f"E{row_num}", {"textFormat": {"foregroundColor": c}})

    # 10) Итого: только если есть строки, где разница реально посчитана
    total_row = end_row + 1
    total_formula = f"=СУММ(F{start_row}:F{end_row})"
    ws.update(
        f"E{total_row}:F{total_row}",
        [["Итого:", total_formula]],
        value_input_option="USER_ENTERED",
    )
    ws.batch_format(
        [
            {"range": f"E{total_row}:F{total_row}", "format": {"textFormat": {"bold": True}}},
            {"range": f"F{total_row}", "format": {"textFormat": {"foregroundColor": {"red": 0, "green": 0, "blue": 0}}}},
            {"range": f"F{total_row}:F{total_row}", "format": {"numberFormat": {"type": "TIME", "pattern": "[h]:mm"}}},
        ]
    )

    # Удаляем старые правила и задаём новые.
    metadata = spreadsheet.fetch_sheet_metadata()
    existing_rules = []
    for sheet in metadata.get("sheets", []):
        props = sheet.get("properties", {})
        if props.get("sheetId") == ws.id:
            existing_rules = sheet.get("conditionalFormats", []) or []
            break

    if existing_rules:
        delete_requests = []
        for idx in reversed(range(len(existing_rules))):
            delete_requests.append({"deleteConditionalFormatRule": {"sheetId": ws.id, "index": idx}})
        spreadsheet.batch_update({"requests": delete_requests})

    rules = [
        # Табель (D) vs Факт (B)
        {
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": [
                        {
                            "sheetId": ws.id,
                            "startRowIndex": start_row - 1,
                            "endRowIndex": end_row,
                            "startColumnIndex": 3,
                            "endColumnIndex": 4,
                        }
                    ],
                    "booleanRule": {
                        "condition": {"type": "CUSTOM_FORMULA", "values": [{"userEnteredValue": f"=D{start_row}<B{start_row}"}]},
                        "format": {"textFormat": {"foregroundColor": _color_red()}},
                    },
                },
                "index": 0,
            }
        },
        {
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": [
                        {
                            "sheetId": ws.id,
                            "startRowIndex": start_row - 1,
                            "endRowIndex": end_row,
                            "startColumnIndex": 3,
                            "endColumnIndex": 4,
                        }
                    ],
                    "booleanRule": {
                        "condition": {"type": "CUSTOM_FORMULA", "values": [{"userEnteredValue": f"=D{start_row}>B{start_row}"}]},
                        "format": {"textFormat": {"foregroundColor": _color_green()}},
                    },
                },
                "index": 1,
            }
        },
        {
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": [
                        {
                            "sheetId": ws.id,
                            "startRowIndex": start_row - 1,
                            "endRowIndex": end_row,
                            "startColumnIndex": 3,
                            "endColumnIndex": 4,
                        }
                    ],
                    "booleanRule": {
                        "condition": {"type": "CUSTOM_FORMULA", "values": [{"userEnteredValue": f"=D{start_row}=B{start_row}"}]},
                        "format": {"textFormat": {"foregroundColor": {"red": 0, "green": 0, "blue": 0}}},
                    },
                },
                "index": 2,
            }
        },
        # Табель (E) vs Факт (C)
        {
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": [
                        {
                            "sheetId": ws.id,
                            "startRowIndex": start_row - 1,
                            "endRowIndex": end_row,
                            "startColumnIndex": 4,
                            "endColumnIndex": 5,
                        }
                    ],
                    "booleanRule": {
                        "condition": {"type": "CUSTOM_FORMULA", "values": [{"userEnteredValue": f"=E{start_row}<C{start_row}"}]},
                        "format": {"textFormat": {"foregroundColor": _color_red()}},
                    },
                },
                "index": 3,
            }
        },
        {
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": [
                        {
                            "sheetId": ws.id,
                            "startRowIndex": start_row - 1,
                            "endRowIndex": end_row,
                            "startColumnIndex": 4,
                            "endColumnIndex": 5,
                        }
                    ],
                    "booleanRule": {
                        "condition": {"type": "CUSTOM_FORMULA", "values": [{"userEnteredValue": f"=E{start_row}>C{start_row}"}]},
                        "format": {"textFormat": {"foregroundColor": _color_green()}},
                    },
                },
                "index": 4,
            }
        },
        {
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": [
                        {
                            "sheetId": ws.id,
                            "startRowIndex": start_row - 1,
                            "endRowIndex": end_row,
                            "startColumnIndex": 4,
                            "endColumnIndex": 5,
                        }
                    ],
                    "booleanRule": {
                        "condition": {"type": "CUSTOM_FORMULA", "values": [{"userEnteredValue": f"=E{start_row}=C{start_row}"}]},
                        "format": {"textFormat": {"foregroundColor": {"red": 0, "green": 0, "blue": 0}}},
                    },
                },
                "index": 5,
            }
        },
        # Разница (F)
        {
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": [
                        {
                            "sheetId": ws.id,
                            "startRowIndex": start_row - 1,
                            "endRowIndex": end_row,
                            "startColumnIndex": 5,
                            "endColumnIndex": 6,
                        }
                    ],
                    "booleanRule": {
                        "condition": {"type": "CUSTOM_FORMULA", "values": [{"userEnteredValue": f"=F{start_row}<0"}]},
                        "format": {"textFormat": {"foregroundColor": _color_red()}},
                    },
                },
                "index": 6,
            }
        },
        {
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": [
                        {
                            "sheetId": ws.id,
                            "startRowIndex": start_row - 1,
                            "endRowIndex": end_row,
                            "startColumnIndex": 5,
                            "endColumnIndex": 6,
                        }
                    ],
                    "booleanRule": {
                        "condition": {"type": "CUSTOM_FORMULA", "values": [{"userEnteredValue": f"=F{start_row}>0"}]},
                        "format": {"textFormat": {"foregroundColor": _color_green()}},
                    },
                },
                "index": 7,
            }
        },
        {
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": [
                        {
                            "sheetId": ws.id,
                            "startRowIndex": start_row - 1,
                            "endRowIndex": end_row,
                            "startColumnIndex": 5,
                            "endColumnIndex": 6,
                        }
                    ],
                    "booleanRule": {
                        "condition": {"type": "CUSTOM_FORMULA", "values": [{"userEnteredValue": f"=F{start_row}=0"}]},
                        "format": {"textFormat": {"foregroundColor": {"red": 0, "green": 0, "blue": 0}}},
                    },
                },
                "index": 8,
            }
        },
    ]
    spreadsheet.batch_update({"requests": rules})

    print(f"✅ Лист '{changes_title}' обновлён. Строк: {len(changes_rows)}")
    return True
