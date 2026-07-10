from __future__ import annotations

import logging
from typing import Any

from ..parsing.common import _round2, _round_row_hours
from ..parsing.contract import _resolve_contracted_hours, load_contract_file_index
from .summary import split_emp_agency_rows, total_paid_hours_from_rows
from .types import PayrollResult

logger = logging.getLogger(__name__)

_HOLIDAY_PAY_FACTOR = 0.1207
HOLIDAY_PAY_FACTOR = _HOLIDAY_PAY_FACTOR


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
