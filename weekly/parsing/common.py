from __future__ import annotations

from datetime import datetime
from typing import Any

import pandas as pd

# ClockRite "Paid Hours (Inc Absence) Summary": Pay ID column B (0-based 1); annual
# hours duplicated at Excel H and L (0-based 7 and 11). See payroll fix plan.
_CLOCKRITE_PAY_ID_COL = 1
_CLOCKRITE_SAGE_HEADER_COL = 3
_CLOCKRITE_ANNUAL_H_COL = 7
_CLOCKRITE_ANNUAL_L_COL = 11
# ClockRite Paid Hours summary: report date in Excel A5 (0-based row 4, col 0).
_CLOCKRITE_PROCESSING_DATE_ROW = 4
_CLOCKRITE_PROCESSING_DATE_COL = 0

_HOUR_DECIMAL_FIELDS = frozenset(
    {
        "BasicHours",
        "MonFriOvertime",
        "SatSunOvertime",
        "AnnualHoliday",
        "TotalPaidHours",
        "ContractedHours",
        "Overtime",
        "ExtraHours",
        "AdditionalHolidayPay",
    }
)


def _normalize_header(text: str) -> str:
    return "".join(ch.lower() for ch in (text or "") if ch.isalnum())


def _to_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    if text.lower() == "nan":
        return ""
    return text


def _parse_decimal(value: Any) -> float:
    text = _to_text(value).replace(",", "")
    if not text:
        return 0.0
    try:
        return float(text)
    except ValueError:
        return 0.0


def _round2(value: float) -> float:
    return round(float(value), 2)


def _round_row_hours(row: dict[str, Any]) -> None:
    for key in _HOUR_DECIMAL_FIELDS:
        if key in row and row[key] is not None:
            row[key] = _round2(row[key])


def _parse_int(value: Any) -> int | None:
    text = _to_text(value).replace(",", "")
    if not text:
        return None
    try:
        return int(float(text))
    except ValueError:
        return None


def _row_contains_date_range(row: list[str]) -> bool:
    return any("date range" in cell.lower() for cell in row)


def _is_clockrite_paid_hours_summary_header(header_cells: list[str], pay_id_col: int) -> bool:
    """True when header row matches Paid Hours (Inc Absence) grid (Pay ID in B, Sage in D)."""
    if pay_id_col != _CLOCKRITE_PAY_ID_COL:
        return False
    if len(header_cells) <= _CLOCKRITE_SAGE_HEADER_COL:
        return False
    return _normalize_header(header_cells[_CLOCKRITE_SAGE_HEADER_COL]) == "sage"


def _annual_holiday_clockrite_hl(row: list[str]) -> float:
    """Annual leave from Excel columns H and L; same value — use H, then L if H empty."""
    h = _parse_decimal(row[_CLOCKRITE_ANNUAL_H_COL] if len(row) > _CLOCKRITE_ANNUAL_H_COL else "")
    l = _parse_decimal(row[_CLOCKRITE_ANNUAL_L_COL] if len(row) > _CLOCKRITE_ANNUAL_L_COL else "")
    if h and l and abs(h - l) > 0.001:
        return h
    if h:
        return h
    return l


def _is_category_row(row: list[str], pay_id_col: int) -> bool:
    pay = row[pay_id_col] if pay_id_col < len(row) else ""
    if _parse_int(pay) is not None:
        return False
    c0 = row[0] if row else ""
    return bool(c0 and not c0.lower().startswith("total for"))


def _load_sheet(file_obj: Any) -> pd.DataFrame:
    file_obj.seek(0)
    return pd.read_excel(file_obj, sheet_name=0, header=None, dtype=str)


def _parse_date_cell(text: str) -> datetime | None:
    text = text.strip()
    if not text:
        return None
    candidates = [text]
    if " " in text:
        candidates.append(text.split(" ", 1)[0])
    formats = ("%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y", "%d.%m.%Y")
    for candidate in candidates:
        for fmt in formats:
            try:
                return datetime.strptime(candidate, fmt)
            except ValueError:
                continue
    return None
