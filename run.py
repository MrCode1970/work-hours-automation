import time
import pandas as pd
import gspread
from playwright.sync_api import sync_playwright
from google.oauth2.service_account import Credentials
from datetime import datetime

# ================= НАСТРОЙКИ =================
USERNAME = "322062415"         # Ваш логин (цифры)
PASSWORD = "1105"           # Ваш пароль
GSHEET_ID = "1rlOdKX8ot0wDCT9BSyySnLkRc18KYbUm5cD_LllRUyU"    # Длинный ID из ссылки Google Таблицы
# =============================================

GOOGLE_JSON_FILE = "service_key.json" 

def get_sheet():
    scopes = ['https://www.googleapis.com/auth/spreadsheets']
    creds = Credentials.from_service_account_file(GOOGLE_JSON_FILE, scopes=scopes)
    client = gspread.authorize(creds)
    return client.open_by_key(GSHEET_ID)

def run():
    print("=== Запуск скрипта ===")
    with sync_playwright() as p:
        # headless=False значит браузер будет ВИДИМЫМ
        browser = p.chromium.launch(headless=False) 
        context = browser.new_context(viewport={'width': 1280, 'height': 800})
        page = context.new_page()
        page.set_default_timeout(60000) # Ждем до 60 сек

        try:
            print("1. Захожу на сайт...")
            page.goto("https://ins.ylm.co.il/#/employeeLogin")
            
            print("2. Ввожу данные...")
            page.wait_for_selector("#Username")
            page.fill("#Username", USERNAME)
            page.fill("#YlmCode", PASSWORD)
            
            print("3. Жму кнопку входа...")
            page.click("button[type='submit']")
            
            print("4. Перехожу в раздел отчетов (דוח הנוכחות שלי)...")
            report_button = "button[ng-click='vm.employeeReport();']"
            page.wait_for_selector(report_button, timeout=60000)
            page.click(report_button)
            
            # --- НОВЫЙ БЛОК: ВЫБОР ТЕКУЩЕГО МЕСЯЦА ---
            print("5. Устанавливаю дату на начало текущего месяца...")
            now = datetime.now()
            first_day_current_month = f"01/{now.strftime('%m/%Y')}" # Формат 01/01/2026
            
            # Ждем поле даты
            date_input = "input[ng-model='vm.report.FromDate']"
            page.wait_for_selector(date_input)
            
            # Очищаем и вводим дату
            page.click(date_input)
            page.keyboard.press("Control+A")
            page.keyboard.press("Backspace")
            page.fill(date_input, first_day_current_month)
            page.keyboard.press("Enter") # Чтобы Angular "подхватил" изменение
            
            print(f"   Установлена дата: {first_day_current_month}")
            time.sleep(1)

            print("6. Нажимаю 'הצג נתונים' (Показать данные)...")
            display_button = "button[ng-click='vm.displayReportResult(true)']"
            page.click(display_button)
            
            # Ждем окончания загрузки (пока кнопка снова не станет активной)
            page.wait_for_load_state("networkidle")
            time.sleep(3) 
            # ------------------------------------------
            
            print("7. Жду кнопку Excel...")
            excel_button = "button[ng-click='executeExcelBtn()']"
            page.wait_for_selector(excel_button, timeout=60000)
            
            print("8. Скачиваю файл...")
            with page.expect_download() as download_info:
                page.click(excel_button)
            
            download = download_info.value
            path = "local_data.xlsx"
            download.save_as(path)
            print("   Файл успешно скачан!")
            
        except Exception as e:
            print(f"!!! ОШИБКА БРАУЗЕРА: {e}")
            page.screenshot(path="debug_screen.png")
            print("   Скриншот экрана сохранен в debug_screen.png")
            browser.close()
            return
# --- ЧАСТЬ 2: Работа с Excel и Гугл Таблицей ---
        print("6. Обработка данных...")
        try:
            df = pd.read_excel(path)
            df_clean = df[['תאריך', 'כניסה', 'יציאה']].dropna()
        except Exception as e:
            print(f"Ошибка чтения Excel: {e}")
            return

        sh = get_sheet()
        now = datetime.now()
        sheet_name = f"{now.month}.{now.strftime('%y')}" 
        
        try:
            worksheet = sh.worksheet(sheet_name)
            print(f"   Лист '{sheet_name}' найден.")
        except:
            print(f"!!! ЛИСТ '{sheet_name}' НЕ НАЙДЕН. В таблице есть: {[s.title for s in sh.worksheets()]}")
            return

        all_values = worksheet.get_all_values()
        updates_count = 0
        print(f"   Начинаю сверку. Всего строк в Excel: {len(df_clean)}")
        
        for index, row in df_clean.iterrows():
            # Извлекаем дату из Excel
            date_raw = str(row['תאריך']).split()[0]
            
            # Подготовка форматов
            try:
                d = pd.to_datetime(date_raw, dayfirst=True)
                variants = [
                    d.strftime('%d/%m/%Y'),       # 01/01/2026
                    d.strftime('%-d/%-m/%Y'),     # 1/1/2026
                    d.strftime('%d.%m.%Y'),       # 01.01.2026
                    d.strftime('%-d.%-m.%y'),     # 1.1.26
                    date_raw                      # Оригинал
                ]
            except:
                variants = [date_raw]

            entry = str(row['כניסה'])[:5]
            exit_ = str(row['יציאה'])[:5]

            found = False
            # Проверяем каждую строку в Google Таблице
            for i, sheet_row in enumerate(all_values):
                if len(sheet_row) > 1:
                    cell_val = str(sheet_row[1]).strip() # Значение в колонке B
                    
                    if not cell_val: continue

                    # Сравниваем
                    if any(v == cell_val for v in variants):
                        worksheet.update_cell(i + 1, 3, entry) # Колонки C и D
                        worksheet.update_cell(i + 1, 4, exit_)
                        updates_count += 1
                        print(f"   [OK] Найдено совпадение: Excel({date_raw}) == GSheet({cell_val})")
                        found = True
                        break
            
            if not found:
                # Если за это число есть время в Excel, но дата не найдена в GSheet
                if entry != "nan" and entry != "00:00":
                    print(f"   [!] Не найдена строка для даты: {date_raw}. Варианты поиска были: {variants}")

        print(f"=== ЗАВЕРШЕНО! Обновлено: {updates_count} ===")
if __name__ == "__main__":
    run()