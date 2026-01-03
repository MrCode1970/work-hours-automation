import os
import json
import time
import pandas as pd
import gspread
from playwright.sync_api import sync_playwright
from google.oauth2.service_account import Credentials
from datetime import datetime

# Настройки из секретов GitHub
USERNAME = os.environ["SITE_USERNAME"]
PASSWORD = os.environ["SITE_PASSWORD"]
GSHEET_ID = os.environ["GSHEET_ID"]
GOOGLE_JSON = json.loads(os.environ["GOOGLE_JSON"])

def get_sheet():
    scopes = ['https://www.googleapis.com/auth/spreadsheets']
    creds = Credentials.from_service_account_info(GOOGLE_JSON, scopes=scopes)
    client = gspread.authorize(creds)
    return client.open_by_key(GSHEET_ID)

def run():
    with sync_playwright() as p:
        # 1. Запуск браузера
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(user_agent="Mozilla/5.0")
        page = context.new_page()

        # 2. Логин
        print("Захожу на сайт...")
        page.goto("https://ins.ylm.co.il/#/employeeLogin")
        page.fill("#Username", USERNAME)
        page.fill("#YlmCode", PASSWORD)
        page.click("button:has-text('כניסה')") # Нажимаем вход (обычно там кнопка с таким текстом)
        
        # Ожидание загрузки после логина
        page.wait_for_load_state("networkidle")
        time.sleep(5) 

        # 3. Скачивание файла
        print("Скачиваю Excel...")
        with page.expect_download() as download_info:
            # Ищем кнопку скачивания по ng-click
            page.click("button[ng-click='executeExcelBtn()']")
        
        download = download_info.value
        path = "data.xlsx"
        download.save_as(path)
        browser.close()

        # 4. Обработка данных
        df = pd.read_excel(path)
        # В вашем файле столбцы: תאריך (A), כניסה (D), יציאה (E)
        # Оставляем только нужные данные
        df_clean = df[['תאריך', 'כניסה', 'יציאה']].dropna()

        # 5. Запись в Google Sheets
        sh = get_sheet()
        # Определяем имя листа как Месяц.Год (например, 1.26)
        now = datetime.now()
        sheet_name = f"{now.month}.{now.strftime('%y')}"
        
        try:
            worksheet = sh.worksheet(sheet_name)
        except:
            print(f"Лист {sheet_name} не найден. Проверьте название!")
            return

        # Получаем все данные из таблицы, чтобы найти нужные строки
        all_values = worksheet.get_all_values()
        
        for index, row in df_clean.iterrows():
            date_str = row['תאריך'] # Формат обычно DD/MM/YYYY
            entry_time = str(row['כניסה'])[:5] # Обрезаем до HH:MM
            exit_time = str(row['יציאה'])[:5]

            # Ищем строку с такой датой в столбце B (индекс 1)
            for i, sheet_row in enumerate(all_values):
                if date_str in sheet_row[1]: # Если дата совпала
                    row_num = i + 1
                    # Обновляем ячейки Вход (C) и Выход (D)
                    worksheet.update_cell(row_num, 3, entry_time) # Столбец C
                    worksheet.update_cell(row_num, 4, exit_time)  # Столбец D
                    print(f"Обновлено: {date_str}")
                    break

run()
