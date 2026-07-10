from __future__ import annotations

from typing import Any

from .common import (
    _CLOCKRITE_PROCESSING_DATE_COL,
    _CLOCKRITE_PROCESSING_DATE_ROW,
    _annual_holiday_clockrite_hl,
    _is_category_row,
    _is_clockrite_paid_hours_summary_header,
    _load_sheet,
    _normalize_header,
    _parse_date_cell,
    _parse_decimal,
    _parse_int,
    _round_row_hours,
    _row_contains_date_range,
    _to_text,
)


def parse_processing_date(file_obj: Any) -> str | None:
    """Report date from ClockRite employee hours file cell A5, as DD.MM.YYYY."""
    df = _load_sheet(file_obj)
    if df.shape[0] <= _CLOCKRITE_PROCESSING_DATE_ROW:
        return None
    raw = df.iat[_CLOCKRITE_PROCESSING_DATE_ROW, _CLOCKRITE_PROCESSING_DATE_COL]
    parsed = _parse_date_cell(_to_text(raw))
    if parsed is None:
        return None
    return parsed.strftime("%d.%m.%Y")


def parse_employee_hours(file_obj: Any) -> list[dict[str, Any]]:
    df = _load_sheet(file_obj)
    rows = [[_to_text(v) for v in rec] for rec in df.values.tolist()]
    if not rows:
        return []

    header_row = -1
    pay_id_col = -1
    scan_limit = min(len(rows), 40)
    for r in range(scan_limit):
        for c, cell in enumerate(rows[r]):
            if _normalize_header(cell) == "payid":
                header_row = r
                pay_id_col = c
                break
        if header_row >= 0:
            break

    if header_row < 0 or pay_id_col < 0:
        return _parse_employee_legacy(rows)

    header_cells = rows[header_row]
    clockrite_grid = _is_clockrite_paid_hours_summary_header(header_cells, pay_id_col)

    name_col = pay_id_col - 1 if pay_id_col > 0 else 0
    basic_col = pay_id_col + 1
    if clockrite_grid:
        # Employee totals: basic C, MF OT D, SS OT E (header may label D/E as Sage / Hrs @ 1 / @ 2).
        mon_fri_col = pay_id_col + 2
        sat_sun_col = pay_id_col + 3
        annual_col = -1
    else:
        mon_fri_col = pay_id_col + 2
        sat_sun_col = pay_id_col + 3
        annual_col = pay_id_col + 4

    category = ""
    result: list[dict[str, Any]] = []
    for r in range(header_row + 1, len(rows)):
        row = rows[r]
        if _row_contains_date_range(row):
            break
        first = row[0] if row else ""
        if first.lower().startswith("total for"):
            continue

        pay_text = row[pay_id_col] if pay_id_col < len(row) else ""
        name_text = row[name_col] if name_col < len(row) else ""
        pay_id = _parse_int(pay_text)
        if pay_id is None:
            if _is_category_row(row, pay_id_col):
                category = first
            continue

        if not name_text or name_text.upper() == "U EMPLOYEE" or "total for" in name_text.lower():
            continue

        basic_gross = _parse_decimal(row[basic_col] if basic_col < len(row) else "")
        mon_fri_ot = _parse_decimal(row[mon_fri_col] if mon_fri_col < len(row) else "")
        sat_sun_ot = _parse_decimal(row[sat_sun_col] if sat_sun_col < len(row) else "")
        if clockrite_grid:
            annual_holiday = _annual_holiday_clockrite_hl(row)
        else:
            annual_holiday = _parse_decimal(row[annual_col] if annual_col >= 0 and annual_col < len(row) else "")
        # Gross basic in export includes annual leave; net basic matches legacy path and FORMULAS #2.
        basic_hours = basic_gross - annual_holiday
        total_paid = basic_hours + mon_fri_ot + sat_sun_ot + annual_holiday

        row = {
            "Name": name_text.upper(),
            "Category": category,
            "SageNo": pay_id,
            "BasicHours": basic_hours,
            "MonFriOvertime": mon_fri_ot,
            "SatSunOvertime": sat_sun_ot,
            "AnnualHoliday": annual_holiday,
            "TotalPaidHours": total_paid,
        }
        _round_row_hours(row)
        result.append(row)

    return result


def _parse_employee_legacy(rows: list[list[str]]) -> list[dict[str, Any]]:
    category = ""
    expects_category = True
    result: list[dict[str, Any]] = []
    idx = 0
    while idx + 6 < len(rows):
        row = rows[6 + idx]
        col_a = row[0] if row else ""
        if col_a:
            if expects_category:
                if col_a.lower() != "default":
                    category = col_a
                expects_category = False
            else:
                if any("date range" in cell.lower() for cell in row):
                    break
                name = _to_text(row[0]).upper()
                if name and name != "U EMPLOYEE":
                    basic = _parse_decimal(row[2] if len(row) > 2 else "")
                    mon_fri = _parse_decimal(row[3] if len(row) > 3 else "")
                    sat_sun = _parse_decimal(row[6] if len(row) > 6 else "")
                    annual = _parse_decimal(row[7] if len(row) > 7 else "")
                    basic_after_legacy = basic - annual
                    result.append(
                        {
                            "Name": name,
                            "Category": category,
                            "SageNo": _parse_int(row[1] if len(row) > 1 else "") or 0,
                            "BasicHours": basic_after_legacy,
                            "MonFriOvertime": mon_fri,
                            "SatSunOvertime": sat_sun,
                            "AnnualHoliday": annual,
                            "TotalPaidHours": basic_after_legacy + mon_fri + sat_sun + annual,
                        }
                    )
        else:
            expects_category = True
            idx += 1
        idx += 1
    return result
