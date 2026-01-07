import os
import json
import time
import pandas as pd
import gspread
from playwright.sync_api import sync_playwright
from google.oauth2.service_account import Credentials
from datetime import datetime

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
        # Используем эмуляцию реального устройства
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(
            viewport={'width': 1920, 'height': 1080},
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        )
        page = context.new_page()
        
        print("Перехожу на страницу логина...")
        try:
            # Ждем только базовой загрузки DOM, а не всех скриптов
            page.goto("https://ins.ylm.co.il/#/employeeLogin", wait_until="commit", timeout=90000)
            
            # Явно ждем появления поля ввода
            print("Ожидаю поле ввода #Username...")
            page.wait_for_selector("#Username", timeout=60000)
            
            page.fill("#Username", USERNAME)
            time.sleep(1)
            page.fill("#YlmCode", PASSWORD)
            time.sleep(1)
            
            print("Нажимаю вход...")
            page.click("button[type='submit']")
            
            print("Ожидаю загрузки личного кабинета (кнопки Excel)...")
            page.wait_for_selector("button[ng-click='executeExcelBtn()']", timeout=90000)

            print("Начинаю скачивание...")
            with page.expect_download() as download_info:
                page.click("button[ng-click='executeExcelBtn()']")
            
            download = download_info.value
            path = "data.xlsx"
            download.save_as(path)
            print("Файл успешно сохранен.")
            
        except Exception as e:
            print(f"Произошла ошибка в браузере: {e}")
            # Делаем скриншот для отладки, если что-то пошло не так
            page.screenshot(path="error_screen.png")
            print("Скриншот ошибки сохранен как error_screen.png")
            browser.close()
            return

        browser.close()

        # --- Обработка данных ---
        print("Читаю Excel и обновляю Google Sheets...")
        df = pd.read_excel(path)
        df_clean = df[['תאריך', 'כניסה', 'יציאה']].dropna()

        sh = get_sheet()
        now = datetime.now()
        sheet_name = f"{now.month}.{now.strftime('%y')}"
        
        try:
            worksheet = sh.worksheet(sheet_name)
        except:
            print(f"Лист {sheet_name} не найден!")
            return

        all_values = worksheet.get_all_values()
        
        for index, row in df_clean.iterrows():
            date_str = str(row['תאריך']).split()[0] # На случай если там есть время
            # Исправляем формат даты (в Excel 2025-12-01, в таблице может быть 01/12/2025)
            # Если в таблице даты через '/', конвертируем:
            try:
                date_obj = pd.to_datetime(date_str)
                formatted_date = date_obj.strftime('%d/%m/%Y')
            except:
                formatted_date = date_str

            entry_time = row['כניסה'].strftime('%H:%M:%S') if hasattr(row['כניסה'], 'strftime') else str(row['כניסה'])[:8]
            exit_time = row['יציאה'].strftime('%H:%M:%S') if hasattr(row['יציאה'], 'strftime') else str(row['יציאה'])[:8]

            for i, sheet_row in enumerate(all_values):
                if len(sheet_row) > 1 and (formatted_date in sheet_row[1] or date_str in sheet_row[1]):
                    row_num = i + 1
                    worksheet.update_cell(row_num, 3, entry_time)
                    worksheet.update_cell(row_num, 4, exit_time)
                    print(f"Обновлено: {formatted_date}")
                    break
        print("Готово!")

if __name__ == "__main__":
    run()
