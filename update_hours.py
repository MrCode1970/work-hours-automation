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
        # 1. Запуск браузера с увеличенным таймаутом
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36"
        )
        page = context.new_page()
        page.set_default_timeout(60000) # Увеличиваем до 60 секунд

        # 2. Логин
        print("Захожу на сайт...")
        try:
            page.goto("https://ins.ylm.co.il/#/employeeLogin", wait_until="networkidle")
        except Exception as e:
            print(f"Сайт грузится долго, но пробуем продолжать... {e}")

        page.fill("#Username", USERNAME)
        time.sleep(1)
        page.fill("#YlmCode", PASSWORD)
        time.sleep(1)
        
        print("Нажимаю кнопку входа...")
        # Кликаем по кнопке submit
        page.click("button[type='submit']") 
        
        # Ждем появления кнопки скачивания (это подтвердит успешный вход)
        print("Ожидаю загрузки личного кабинета...")
        page.wait_for_selector("button[ng-click='executeExcelBtn()']", timeout=60000)

        # 3. Скачивание файла
        print("Скачиваю Excel...")
        with page.expect_download() as download_info:
            page.click("button[ng-click='executeExcelBtn()']")
        
        download = download_info.value
        path = "data.xlsx"
        download.save_as(path)
        print("Файл скачан успешно.")
        browser.close()

        # 4. Обработка данных
        df = pd.read_excel(path)
        # Удаляем пустые строки и берем нужные колонки
        df_clean = df[['תאריך', 'כניסה', 'יציאה']].dropna()

        # 5. Запись в Google Sheets
        sh = get_sheet()
        now = datetime.now()
        # Лист формата 1.26
        sheet_name = f"{now.month}.{now.strftime('%y')}"
        
        try:
            worksheet = sh.worksheet(sheet_name)
        except:
            print(f"Лист {sheet_name} не найден!")
            return

        all_values = worksheet.get_all_values()
        
        updates = []
        for index, row in df_clean.iterrows():
            date_str = row['תאריך']
            # Время может быть объектом datetime.time, приводим к строке HH:MM
            entry_time = row['כניסה'].strftime('%H:%M') if hasattr(row['כניסה'], 'strftime') else str(row['כניסה'])[:5]
            exit_time = row['יציאה'].strftime('%H:%M') if hasattr(row['יציאה'], 'strftime') else str(row['יציאה'])[:5]

            for i, sheet_row in enumerate(all_values):
                if len(sheet_row) > 1 and date_str in sheet_row[1]: # Дата в колонке B (индекс 1)
                    row_num = i + 1
                    # Готовим данные для обновления (C и D)
                    worksheet.update_cell(row_num, 3, entry_time)
                    worksheet.update_cell(row_num, 4, exit_time)
                    print(f"Обновлена дата: {date_str}")
                    break
        print("Все данные синхронизированы.")

if __name__ == "__main__":
    run()
