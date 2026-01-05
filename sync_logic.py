import pandas as pd


def sync_from_excel(worksheet, excel_path: str) -> None:
    """
    Читает Excel-файл и обновляет Google Sheet.
    Логика соответствует текущему рабочему run.py:
    - берём колонки תאריך / כניסה / יציאה
    - ищем совпадение даты в столбце B
    - записываем вход/выход в колонки C/D
    """
    df = pd.read_excel(excel_path)
    df_clean = df[["תאריך", "כניסה", "יציאה"]].dropna()

    all_values = worksheet.get_all_values()

    for _, row in df_clean.iterrows():
        raw_date = str(row["תאריך"]).split()[0]

        # Варианты формата даты — как в твоём run.py
        try:
            d = pd.to_datetime(raw_date, dayfirst=True)
            date_variants = [
                d.strftime("%d/%m/%Y"),
                d.strftime("%-d/%-m/%Y"),
                raw_date,
                d.strftime("%d.%m.%Y"),
                d.strftime("%-d.%-m.%y"),
            ]
        except Exception:
            date_variants = [raw_date]

        entry = str(row["כניסה"])[:5]
        exit_ = str(row["יציאה"])[:5]

        found = False
        for i, sheet_row in enumerate(all_values):
            if len(sheet_row) > 1:
                sheet_date = str(sheet_row[1]).strip()
                if any(v == sheet_date for v in date_variants):
                    worksheet.update_cell(i + 1, 3, entry)
                    worksheet.update_cell(i + 1, 4, exit_)
                    found = True
                    break

        if not found and entry not in ("nan", "00:00"):
            print(f"[!] Не найдена строка для даты: {raw_date}. Варианты: {date_variants}")
