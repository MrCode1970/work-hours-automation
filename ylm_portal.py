import os
import random
import time
from datetime import datetime
from typing import Callable, Iterable
from playwright.sync_api import expect, sync_playwright

from ylm_actions import build_actions


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
            run_actions(page, build_actions(site_username, site_password))

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
            sleep_action_delay()

            attempts = 3
            last_error = None
            for attempt in range(1, attempts + 1):
                print(f"⬇️ Попытка скачивания {attempt}/{attempts}")
                try:
                    page.reload(wait_until="networkidle")
                    page.wait_for_selector(excel_button)
                    sleep_action_delay()
                    with page.expect_download(timeout=60000) as download_info:
                        page.click(excel_button)
                    download = download_info.value
                    download.save_as(excel_path)

                    if not os.path.exists(excel_path) or os.path.getsize(excel_path) <= 0:
                        raise RuntimeError("Скачанный файл отсутствует или пустой")

                    print(f"✅ Скачивание успешно: {excel_path}")
                    return excel_path
                except Exception as exc:
                    last_error = exc
                    print(f"⚠️ Скачивание не удалось: {exc}")
                    if attempt < attempts:
                        print("🔄 Перезагружаю страницу и пробую снова...")
                        page.reload(wait_until="networkidle")
                        page.wait_for_selector(excel_button)
                        locator = page.locator(excel_button)
                        locator.scroll_into_view_if_needed()
                        locator.wait_for(state="visible", timeout=30000)
                        expect(locator).to_be_enabled(timeout=30000)
                        sleep_action_delay()
                        continue
                    break

            raise RuntimeError(
                f"Не удалось скачать Excel за {attempts} попытки. Последняя ошибка: {last_error}"
            )

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


def _parse_delay(raw: str) -> tuple[float, float]:
    raw = (raw or "").strip()
    if not raw:
        return 0.0, 0.0
    if "-" in raw:
        lo, hi = raw.split("-", 1)
        return float(lo), float(hi)
    val = float(raw)
    return val, val


def _get_action_delay() -> tuple[float, float]:
    """
    ACTION_DELAY=3-5 or ACTION_DELAY=2
    """
    return _parse_delay(os.getenv("ACTION_DELAY", "0"))


def sleep_action_delay() -> None:
    lo, hi = _get_action_delay()
    if hi <= 0:
        return
    if hi < lo:
        lo, hi = hi, lo
    delay = random.uniform(lo, hi)
    time.sleep(delay)


Step = Callable[[], None]


def run_steps(steps: Iterable[Step]) -> None:
    for step in steps:
        step()
        sleep_action_delay()


def run_actions(page, actions: Iterable[dict]) -> None:
    def _step(action: dict) -> Step:
        kind = action["type"]
        if kind == "goto":
            return lambda: page.goto(action["url"], wait_until=action.get("wait_until", "domcontentloaded"))
        if kind == "wait":
            return lambda: page.wait_for_selector(action["selector"], timeout=action.get("timeout", 60000))
        if kind == "fill":
            return lambda: page.fill(action["selector"], action["value"])
        if kind == "click":
            return lambda: page.click(action["selector"])
        if kind == "reload":
            return lambda: page.reload(wait_until=action.get("wait_until", "networkidle"))
        raise ValueError(f"Unknown action type: {kind}")

    run_steps(_step(a) for a in actions)
