import time
from datetime import datetime
from playwright.sync_api import sync_playwright


def download_excel(site_username: str, site_password: str, excel_path: str = "local_data.xlsx", headless: bool = False) -> str:
    """
    Логин на ylm.co.il и скачивание Excel отчёта за текущий месяц.
    Возвращает путь к сохранённому файлу excel_path.
    """
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=headless)
        context = browser.new_context()
        page = context.new_page()

        page.set_default_timeout(120000)
        page.set_default_navigation_timeout(120000)

        # Trace — суперполезно в CI
        context.tracing.start(screenshots=True, snapshots=True, sources=True)

        try:
            url = "https://ins.ylm.co.il/#/employeeLogin"
            page.goto(url, wait_until="domcontentloaded")

            page.fill("#Username", site_username)
            page.fill("#YlmCode", site_password)
            page.click("button[type='submit']")

            report_button = "button[ng-click='vm.employeeReport();']"
            page.wait_for_selector(report_button)
            page.click(report_button)

            now = datetime.now()
            first_day = f"01/{now.strftime('%m/%Y')}"
            date_input = "input[ng-model='vm.report.FromDate']"
            page.wait_for_selector(date_input)

            page.click(date_input)
            page.keyboard.press("Control+A")
            page.keyboard.press("Backspace")
            page.fill(date_input, first_day)
            page.keyboard.press("Enter")
            time.sleep(1)

            display_button = "button[ng-click='vm.displayReportResult(true)']"
            page.click(display_button)
            page.wait_for_load_state("networkidle")
            time.sleep(2)

            excel_button = "button[ng-click='executeExcelBtn()']"
            page.wait_for_selector(excel_button)

            with page.expect_download() as download_info:
                page.click(excel_button)
            download = download_info.value
            download.save_as(excel_path)

            return excel_path

        except Exception:
            try:
                page.screenshot(path="debug_screen.png", full_page=True)
            except Exception:
                pass
            try:
                html = page.content()
                with open("debug_page.html", "w", encoding="utf-8") as f:
                    f.write(html)
            except Exception:
                pass
            raise

        finally:
            # trace пытаемся сохранить всегда
            try:
                context.tracing.stop(path="debug_trace.zip")
            except Exception:
                pass
            browser.close()
