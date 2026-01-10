from __future__ import annotations

from typing import Any


def build_actions(site_username: str, site_password: str, first_day: str) -> list[dict[str, Any]]:
    report_button = "button[ng-click='vm.employeeReport();']"
    date_input = "input[ng-model='vm.report.FromDate']"
    display_button = "button[ng-click='vm.displayReportResult(true)']"
    excel_button = "button[ng-click='executeExcelBtn()']"

    return [
        {"type": "goto", "url": "https://ins.ylm.co.il/#/employeeLogin", "wait_until": "domcontentloaded"},
        {"type": "wait", "selector": "#Username", "timeout": 60000},
        {"type": "fill", "selector": "#Username", "value": site_username},
        {"type": "fill", "selector": "#YlmCode", "value": site_password},
        {"type": "click", "selector": "button[type='submit']"},
        {"type": "wait", "selector": report_button},
        {"type": "click", "selector": report_button},
        {"type": "wait", "selector": date_input},
        {"type": "click", "selector": date_input},
        {"type": "press", "key": "Control+A"},
        {"type": "press", "key": "Backspace"},
        {"type": "fill", "selector": date_input, "value": first_day},
        {"type": "press", "key": "Enter"},
        {"type": "sleep", "seconds": 1},
        {"type": "click", "selector": display_button},
        {"type": "wait_load_state", "state": "networkidle"},
        {"type": "sleep", "seconds": 2},
        {"type": "wait", "selector": excel_button},
        {"type": "sleep", "seconds": 3},
        {
            "type": "download",
            "selector": excel_button,
            "attempts": 3,
            "reload_before_click": True,
        },
    ]
