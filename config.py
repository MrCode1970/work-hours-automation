import os
from typing import Optional


def get_env(name: str, default: Optional[str] = None) -> str:
    val = os.getenv(name, default)
    if val is None or str(val).strip() == "":
        raise RuntimeError(f"Не задана переменная окружения: {name}")
    return str(val).strip()


def get_headless() -> bool:
    """
    HEADLESS=1/true/yes -> headless True, иначе False.
    По умолчанию False (удобно локально).
    """
    raw = os.getenv("HEADLESS", "0").strip().lower()
    return raw in ("1", "true", "yes", "y", "on")


def get_bool_env(name: str, default: str = "0") -> bool:
    """
    Универсальный bool из env.
    """
    raw = os.getenv(name, default).strip().lower()
    return raw in ("1", "true", "yes", "y", "on")


def load_config() -> dict:
    """
    Единая точка получения конфигурации.
    """
    return {
        "SITE_USERNAME": get_env("SITE_USERNAME"),
        "SITE_PASSWORD": get_env("SITE_PASSWORD"),
        "GSHEET_ID": get_env("GSHEET_ID"),
        # Локально: файл service_key.json в корне.
        # В GitHub Actions можно создавать этот файл из секретов.
        "GOOGLE_JSON_FILE": os.getenv("GOOGLE_JSON_FILE", "service_key.json").strip(),
        "HEADLESS": get_headless(),
        # Имя временного Excel-файла
        "EXCEL_PATH": os.getenv("EXCEL_PATH", "local_data.xlsx").strip(),
        # Явный источник файла (pdf/xlsx). Если задан, используется вместо скачивания.
        "SOURCE_FILE": os.getenv("SOURCE_FILE", "").strip(),
        # Если 1/true/yes — НЕ открывать сайт, использовать уже существующий Excel
        "SKIP_DOWNLOAD": get_bool_env("SKIP_DOWNLOAD", "0"),
        # Если 1/true/yes — ручной режим: логин и скачивание делает пользователь
        "MANUAL_PORTAL": get_bool_env("MANUAL_PORTAL", "0"),
        # Таймаут ожидания скачивания в ручном режиме (мс). 0 = без таймаута.
        "MANUAL_DOWNLOAD_TIMEOUT_MS": int(os.getenv("MANUAL_DOWNLOAD_TIMEOUT_MS", "0").strip() or "0"),
        # Если 1/true/yes — открыть портал в мобильном UI для ручных действий и выйти.
        "MOBILE_UI": get_bool_env("MOBILE_UI", "0"),
        # Device preset для мобильного UI (Playwright). Пример: "iPhone 14 Pro Max".
        "MOBILE_DEVICE": os.getenv("MOBILE_DEVICE", "iPhone 14 Pro Max").strip(),
    }
