from __future__ import annotations

from typing import Any


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
