from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


@dataclass
class RawAttendance:
    meta: dict[str, Any] = field(default_factory=dict)
    rows: list[dict[str, str]] = field(default_factory=list)

    def as_dict(self) -> dict[str, Any]:
        return {"meta": self.meta, "rows": self.rows}
