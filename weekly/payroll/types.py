from __future__ import annotations

from dataclasses import dataclass
from typing import Any


@dataclass
class PayrollResult:
    rows: list[dict[str, Any]]
    agency_rows: list[dict[str, Any]]
    gazebo_rows: list[dict[str, Any]]
    total_paid_hours: float
