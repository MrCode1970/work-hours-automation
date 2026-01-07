from __future__ import annotations

from datetime import datetime
import pandas as pd


def _norm_time(v) -> str:
    """
    Нормализует время к HH:MM.
    Принимает строки вида "7:00", "07:00", "07:00:00", а также NaN/None.
    """
    if v is None:
        return ""
    s = str(v).strip()
    if s == "" or s.lower() == "nan":
        return ""
    s = s[:8]  # на случай "07:00:00"
    if ":" not in s:
        return ""

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
        return ""


def _time_to_minutes(t: str) -> int:
    if not t or ":" not in t:
        return 0
    h, m = t.split(":")[:2]
    return int(h) * 60 + int(m)


def _minutes_to_hhmm_signed(m: int) -> str:
    """
    Возвращает строку вида -2:00 или +0:30
    """
    sign = "-" if m < 0 else "+"
    m = abs(m)
    return f"{sign}{m // 60}:{m % 60:02d}"


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


def build_changes_sheet(spreadsheet, base_ws, sheet_name: str, excel_path: str) -> None:
    """
    Создаёт/пересоздаёт лист "Изменения M.YY" (состояние расхождений).
    Если расхождений нет — лист удаляется (или не создаётся).

    Дополнительно:
    - если в основном листе (M.YY) вход/выход пустые, а на сайте (Excel) есть значения,
      то мы ДОПОЛНЯЕМ пустые ячейки (C/D) значениями с сайта.
      Это НЕ считается "изменением", это просто заполнение отсутствующих данных.

    Столбцы листа изменений строго:
    Дата | Вход (мой) | Выход (мой) | Вход (сайт) | Выход (сайт) | Разница
    """

    changes_title = f"Изменения {sheet_name}"

    # 1) Считаем Excel (сайт)
    df = pd.read_excel(excel_path)
    if not all(c in df.columns for c in ["תאריך", "כניסה", "יציאה"]):
        raise RuntimeError("Excel не содержит ожидаемые колонки: תאריך, כניסה, יציאה")
    df = df[["תאריך", "כניסה", "יציאה"]].dropna(subset=["תאריך"])

    # 2) Считаем базовую таблицу (твои часы — эталон)
    base_values = base_ws.get_all_values()

    # Индекс по дате из base:
    # date_str -> (row_num, my_in, my_out)
    # В base дата в столбце B (index 1), вход/выход в C/D (2/3)
    base_by_date: dict[str, tuple[int, str, str]] = {}
    for idx, row in enumerate(base_values):
        if len(row) < 2:
            continue
        date_cell = str(row[1]).strip()
        if not date_cell:
            continue

        my_in = _norm_time(row[2] if len(row) > 2 else "")
        my_out = _norm_time(row[3] if len(row) > 3 else "")
        row_num = idx + 1  # 1-based for Sheets API

        base_by_date[date_cell] = (row_num, my_in, my_out)

    # 3) Собираем изменения
    # Каждая строка: [date, my_in, my_out, site_in, site_out, diff_str, diff_minutes, cmp_in, cmp_out]
    changes_rows = []
    base_updates = []

    # Вспомогательно: кэш обновлённых "моих" значений, чтобы сразу сравнивать корректно
    # date_str -> (my_in, my_out)
    updated_my_cache: dict[str, tuple[str, str]] = {}

    for _, r in df.iterrows():
        raw_date = str(r["תאריך"]).split()[0]
        try:
            d = pd.to_datetime(raw_date, dayfirst=True)
            date_variants = [d.strftime("%d.%m.%Y"), d.strftime("%d/%m/%Y")]
        except Exception:
            continue

        site_in = _norm_time(r["כניסה"])
        site_out = _norm_time(r["יציאה"])

        # Найдём дату в основном листе (по одному из вариантов формата)
        base_date = None
        row_num = None
        my_in = ""
        my_out = ""
        for dv in date_variants:
            if dv in base_by_date:
                base_date = dv
                row_num, my_in, my_out = base_by_date[dv]
                break

        if base_date is None or row_num is None:
            # даты нет в твоём листе — это не "изменение"
            continue

        # Если ранее уже дополняли эту дату в этом запуске — используем актуальные значения
        if base_date in updated_my_cache:
            my_in, my_out = updated_my_cache[base_date]

        # 3a) ДОПОЛНЯЕМ пустые "мои" значения из сайта (только если у меня пусто, а у сайта есть)
        changed_base = False
        if my_in == "" and site_in != "":
            base_updates.append({"range": f"C{row_num}", "values": [[site_in]]})
            my_in = site_in
            changed_base = True

        if my_out == "" and site_out != "":
            base_updates.append({"range": f"D{row_num}", "values": [[site_out]]})
            my_out = site_out
            changed_base = True

        if changed_base:
            updated_my_cache[base_date] = (my_in, my_out)

        # 3b) После дополнения проверяем: есть ли РЕАЛЬНОЕ расхождение
        if my_in == site_in and my_out == site_out:
            continue

        # 3c) Сравнение для окраски значений сайта относительно "моих"
        cmp_in = 0
        cmp_out = 0
        if my_in and site_in and my_in != site_in:
            cmp_in = -1 if _time_to_minutes(site_in) < _time_to_minutes(my_in) else 1
        if my_out and site_out and my_out != site_out:
            cmp_out = -1 if _time_to_minutes(site_out) < _time_to_minutes(my_out) else 1

        # 3d) Разницу по дням считаем ТОЛЬКО если у обеих сторон есть вход+выход
        can_calc = bool(my_in and my_out and site_in and site_out)
        if can_calc:
            my_minutes = _time_to_minutes(my_out) - _time_to_minutes(my_in)
            site_minutes = _time_to_minutes(site_out) - _time_to_minutes(site_in)
            diff_minutes = site_minutes - my_minutes
            diff_str = _minutes_to_hhmm_signed(diff_minutes)
        else:
            diff_minutes = None
            diff_str = ""

        changes_rows.append([base_date, my_in, my_out, site_in, site_out, diff_str, diff_minutes, cmp_in, cmp_out])

    # 4) Применяем дозаполнения в основной таблице одним пакетом
    if base_updates:
        try:
            base_ws.batch_update(base_updates)
        except AttributeError:
            for u in base_updates:
                base_ws.update(u["range"], u["values"])

    # 5) Если расхождений нет — удалить лист и выйти
    if not changes_rows:
        _delete_worksheet_if_exists(spreadsheet, changes_title)
        print(f"✅ Расхождений нет — лист '{changes_title}' удалён/не создан.")
        return

    # 6) Пересоздать лист изменений
    _delete_worksheet_if_exists(spreadsheet, changes_title)
    ws = spreadsheet.add_worksheet(title=changes_title, rows=len(changes_rows) + 10, cols=6)

    # 6) A1 и заголовки
    ws.update("A1", [[f"Дата изменений: {datetime.now().strftime('%d.%m.%Y')}"]])

    headers = ["Дата", "Вход (мой)", "Выход (мой)", "Вход (сайт)", "Выход (сайт)", "Разница"]
    ws.update("A3:F3", [headers])
    ws.format("A3:F3", {"textFormat": {"bold": True}})

    # 7) Данные одним блоком
    start_row = 4
    values_block = []
    for rr in changes_rows:
        values_block.append([rr[0], rr[1], rr[2], rr[3], rr[4], rr[5]])

    end_row = start_row + len(values_block) - 1
    ws.update(f"A{start_row}:F{end_row}", values_block)

    # 8) Фон групп (как у тебя по образцу)
    ws.format(f"B3:C{end_row}", {"backgroundColor": _bg_my()})
    ws.format(f"D3:E{end_row}", {"backgroundColor": _bg_site()})

    # 9) Окраска текста: сайт и разница
    total_diff = 0
    any_total = False

    for idx, rr in enumerate(changes_rows):
        row_num = start_row + idx

        diff_minutes = rr[6]
        if diff_minutes is not None:
            any_total = True
            total_diff += diff_minutes
            diff_color = _color_red() if diff_minutes < 0 else _color_green()
            ws.format(f"F{row_num}", {"textFormat": {"foregroundColor": diff_color}})

        cmp_in = rr[7]
        cmp_out = rr[8]
        if cmp_in != 0:
            c = _color_red() if cmp_in < 0 else _color_green()
            ws.format(f"D{row_num}", {"textFormat": {"foregroundColor": c}})
        if cmp_out != 0:
            c = _color_red() if cmp_out < 0 else _color_green()
            ws.format(f"E{row_num}", {"textFormat": {"foregroundColor": c}})

    # 10) Итого: только если есть строки, где разница реально посчитана
    total_row = end_row + 2
    total_str = _minutes_to_hhmm_signed(total_diff) if any_total else ""
    ws.update(f"E{total_row}:F{total_row}", [["Итого:", total_str]])
    ws.format(f"E{total_row}", {"textFormat": {"bold": True}})
    ws.format(f"F{total_row}", {"textFormat": {"bold": True}})

    if any_total:
        total_color = _color_red() if total_diff < 0 else _color_green()
        ws.format(f"F{total_row}", {"textFormat": {"foregroundColor": total_color}})

    print(f"✅ Лист '{changes_title}' обновлён. Строк: {len(changes_rows)}")
