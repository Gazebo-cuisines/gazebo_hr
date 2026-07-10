from __future__ import annotations

import logging
from dataclasses import dataclass
from typing import Any

from .parsing.common import (
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
    _round2,
    _round_row_hours,
    _row_contains_date_range,
    _to_text,
)
from .parsing.contract import (
    ContractAuditResult,
    _resolve_contracted_hours,
    audit_contract_integrity,
    audit_contract_pay_id_coverage,
    load_contract_file_index,
    parse_contracted_hours,
    parse_employee_display_names,
)

logger = logging.getLogger(__name__)


def is_agency_category(category: str) -> bool:
    """Agency rows use ClockRite categories whose name starts with A-."""
    return str(category or "").strip().upper().startswith("A-")


def agency_categories_from_rows(rows: list[dict[str, Any]]) -> list[str]:
    """Distinct category labels from rows that count as agency (A- prefix), sorted."""
    seen: set[str] = set()
    out: list[str] = []
    for row in rows:
        cat = str(row.get("Category", "")).strip()
        if is_agency_category(cat) and cat not in seen:
            seen.add(cat)
            out.append(cat)
    return sorted(out, key=str.upper)


def split_emp_agency_rows(
    employee_rows: list[dict[str, Any]],
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    """Return (gazebo_rows, agency_rows) for EMP / Agency totals."""
    agency_rows = [r for r in employee_rows if is_agency_category(str(r.get("Category", "")))]
    gazebo_rows = [r for r in employee_rows if not is_agency_category(str(r.get("Category", "")))]
    return gazebo_rows, agency_rows


@dataclass
class PayrollResult:
    rows: list[dict[str, Any]]
    agency_rows: list[dict[str, Any]]
    gazebo_rows: list[dict[str, Any]]
    total_paid_hours: float



def total_paid_hours_from_rows(rows: list[dict[str, Any]]) -> float:
    """Sum TotalPaidHours across all employee rows (includes every category, including A- prefix)."""
    return round(sum(float(r.get("TotalPaidHours", 0.0)) for r in rows), 2)


def _work_hours(row: dict[str, Any]) -> float:
    return (
        float(row.get("BasicHours", 0.0) or 0.0)
        + float(row.get("MonFriOvertime", 0.0) or 0.0)
        + float(row.get("SatSunOvertime", 0.0) or 0.0)
    )


def build_staff_summary(rows: list[dict[str, Any]]) -> dict[str, Any]:
    """Headcounts for Day Report cover: agency, gazebo worked, gazebo paid-holiday only."""
    gazebo_rows, agency_rows = split_emp_agency_rows(rows)
    agency_staff = sum(1 for r in agency_rows if float(r.get("TotalPaidHours", 0.0) or 0.0) > 0)
    gazebo_paid_holiday = sum(
        1
        for r in gazebo_rows
        if float(r.get("TotalPaidHours", 0.0) or 0.0) > 0
        and float(r.get("AnnualHoliday", 0.0) or 0.0) > 0
        and _work_hours(r) == 0
    )
    gazebo_staff = sum(1 for r in gazebo_rows if _work_hours(r) > 0)
    return {
        "total_staff": agency_staff + gazebo_staff + gazebo_paid_holiday,
        "agency_staff": agency_staff,
        "gazebo_staff": gazebo_staff,
        "gazebo_paid_holiday": gazebo_paid_holiday,
        "total_paid_hours": total_paid_hours_from_rows(rows),
    }



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




def calculate_payroll(employee_rows: list[dict[str, Any]], contracted_file_obj: Any) -> PayrollResult:
    contracted_file_obj.seek(0)
    index = load_contract_file_index(contracted_file_obj)

    for row in employee_rows:
        name_upper = str(row["Name"]).upper()
        contracted, match, reason = _resolve_contracted_hours(int(row["SageNo"]), name_upper, index)
        row["ContractHourMatch"] = match
        row["ContractMatchReason"] = reason
        if match == "No":
            logger.debug(
                "Contract match miss: SageNo=%s Name=%s (%s)",
                row.get("SageNo"),
                row.get("Name"),
                reason,
            )
        row["ContractedHours"] = _round2(contracted)
        row["Overtime"] = _round2(max(0.0, float(row["TotalPaidHours"]) - contracted))
        _round_row_hours(row)

    for row in employee_rows:
        full = index.by_sage_name.get(int(row["SageNo"])) or index.by_clock_name.get(str(row["Name"]).upper())
        if full:
            row["Name"] = full

    gazebo_rows, agency_rows = split_emp_agency_rows(employee_rows)
    all_hours = total_paid_hours_from_rows(employee_rows)

    return PayrollResult(
        rows=employee_rows,
        agency_rows=agency_rows,
        gazebo_rows=gazebo_rows,
        total_paid_hours=all_hours,
    )


_HOLIDAY_PAY_FACTOR = 0.1207
HOLIDAY_PAY_FACTOR = _HOLIDAY_PAY_FACTOR


def compute_extra_holiday_pay(
    total_paid_hours: float,
    contracted_hours: float = 0.0,
    *,
    extra_hours: float | None = None,
    additional_holiday_pay: float | None = None,
) -> tuple[float, float]:
    """Weekly report rules: extra = max(0, actual − contracted); holiday = factor × extra."""
    if extra_hours is not None:
        extra = _round2(max(0.0, float(extra_hours)))
    else:
        actual = max(0.0, float(total_paid_hours))
        contracted = max(0.0, float(contracted_hours))
        extra = _round2(max(0.0, actual - contracted))
    if additional_holiday_pay is not None:
        holiday = _round2(max(0.0, float(additional_holiday_pay)))
    else:
        holiday = _round2(max(0.0, extra * _HOLIDAY_PAY_FACTOR))
    return extra, holiday


def calculate_weekly_payroll(employee_rows: list[dict[str, Any]], contracted_file_obj: Any) -> PayrollResult:
    """Like calculate_payroll, with clamped hours, ExtraHours, and AdditionalHolidayPay."""
    result = calculate_payroll(employee_rows, contracted_file_obj)
    for row in result.rows:
        actual = max(0.0, float(row["TotalPaidHours"]))
        contracted = max(0.0, float(row["ContractedHours"]))
        row["TotalPaidHours"] = actual
        row["ContractedHours"] = contracted
        row.pop("Overtime", None)
        extra, holiday = compute_extra_holiday_pay(actual, contracted)
        row["ExtraHours"] = extra
        row["AdditionalHolidayPay"] = holiday
        _round_row_hours(row)
    return PayrollResult(
        rows=result.rows,
        agency_rows=result.agency_rows,
        gazebo_rows=result.gazebo_rows,
        total_paid_hours=total_paid_hours_from_rows(result.rows),
    )



from .excel.workbook import (  # noqa: E402  — facade re-export after defs (avoids circular import)
    _HOUR_BAND_COLS,
    _OVERALL_CATEGORY_ORDER,
    _build_analysis_dataframe,
    _build_grouped_analysis_dataframe,
    _build_overall_analysis_dataframe,
    _overall_category_key,
    _sum_hour_bands,
    build_category_summary_hr_df,
    build_emp_agency_total_df,
    build_excel_bytes,
    build_hours_over_60_df,
    build_overall_category_totals,
)
