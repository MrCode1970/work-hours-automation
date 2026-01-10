from __future__ import annotations

from typing import Any


def build_actions(site_username: str, site_password: str) -> list[dict[str, Any]]:
    return [
        {"type": "goto", "url": "https://ins.ylm.co.il/#/employeeLogin", "wait_until": "domcontentloaded"},
        {"type": "wait", "selector": "#Username", "timeout": 60000},
        {"type": "fill", "selector": "#Username", "value": site_username},
        {"type": "fill", "selector": "#YlmCode", "value": site_password},
        {"type": "click", "selector": "button[type='submit']"},
        {"type": "wait", "selector": "button[ng-click='vm.employeeReport();']"},
        {"type": "click", "selector": "button[ng-click='vm.employeeReport();']"},
    ]
