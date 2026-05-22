from __future__ import annotations

from dataclasses import asdict, dataclass, field
from io import BytesIO
from typing import Any

import pandas as pd
from openpyxl import Workbook


_BAND_KEYS = ("BasicHours", "MonFriOvertime", "SatSunOvertime", "AnnualHoliday", "TotalPaidHours")
_EMP_AGENCY_ROWS = ("EMP", "AGENCY", "TOTAL")


@dataclass
class MonthlyEmployee:
    Name: str
    Category: str
    SageNo: int
    BasicHours: float
    MonFriOvertime: float
    SatSunOvertime: float
    AnnualHoliday: float
    TotalPaidHours: float
    IsHourly: bool = True


@dataclass
class MonthlyEmployeeTotal:
    Category: str
    BasicHours: float
    MonFriOvertime: float
    SatSunOvertime: float
    AnnualHoliday: float
    TotalPaidHours: float


@dataclass
class MonthlyAdjustment:
    Name: str
    Type: str
    Value: float


@dataclass
class MonthlyWeekSummary:
    employees: list[MonthlyEmployee] = field(default_factory=list)
    employee_totals: list[MonthlyEmployeeTotal] = field(default_factory=list)
    adjustments: list[MonthlyAdjustment] = field(default_factory=list)
    start_date: str = ""
    end_date: str = ""
    non_agency_total: float = 0.0
    grouped_totals: dict[str, MonthlyEmployeeTotal] = field(default_factory=dict)
    emp_agency_bands: dict[str, dict[str, float]] = field(default_factory=dict)


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


def _parse_int(value: Any) -> int:
    text = _to_text(value).replace(",", "")
    if not text:
        return 0
    try:
        return int(float(text))
    except ValueError:
        return 0


def _grouped_key(category: str) -> str:
    category = _to_text(category)
    if category.startswith("A-") and len(category) >= 9:
        return f"{category[:4]} {category[5:9]}"
    return category[:4] if len(category) >= 4 else category


def _is_agency_category(category: str) -> bool:
    return _to_text(category).upper().startswith("A-")


def _empty_bands() -> dict[str, float]:
    return {k: 0.0 for k in _BAND_KEYS}


def _compute_emp_agency_bands(employees: list[MonthlyEmployee]) -> dict[str, dict[str, float]]:
    emp = _empty_bands()
    agency = _empty_bands()
    for e in employees:
        target = agency if _is_agency_category(e.Category) else emp
        target["BasicHours"] += e.BasicHours
        target["MonFriOvertime"] += e.MonFriOvertime
        target["SatSunOvertime"] += e.SatSunOvertime
        target["AnnualHoliday"] += e.AnnualHoliday
        target["TotalPaidHours"] += e.TotalPaidHours
    total = {k: emp[k] + agency[k] for k in _BAND_KEYS}
    return {"EMP": emp, "AGENCY": agency, "TOTAL": total}


def _sum_emp_agency_bands(week_bands: list[dict[str, dict[str, float]]]) -> dict[str, dict[str, float]]:
    out = {row: _empty_bands() for row in _EMP_AGENCY_ROWS}
    for bands in week_bands:
        for row in _EMP_AGENCY_ROWS:
            for k in _BAND_KEYS:
                out[row][k] += float(bands.get(row, {}).get(k, 0.0) or 0.0)
    return out


def _employee_totals_from_employees(employees: list[MonthlyEmployee]) -> list[MonthlyEmployeeTotal]:
    by_cat: dict[str, MonthlyEmployeeTotal] = {}
    for e in employees:
        if e.Category not in by_cat:
            by_cat[e.Category] = MonthlyEmployeeTotal(e.Category, 0.0, 0.0, 0.0, 0.0, 0.0)
        t = by_cat[e.Category]
        t.BasicHours += e.BasicHours
        t.MonFriOvertime += e.MonFriOvertime
        t.SatSunOvertime += e.SatSunOvertime
        t.AnnualHoliday += e.AnnualHoliday
        t.TotalPaidHours += e.TotalPaidHours
    return list(by_cat.values())


def _build_grouped_totals(totals: list[MonthlyEmployeeTotal]) -> dict[str, MonthlyEmployeeTotal]:
    grouped: dict[str, MonthlyEmployeeTotal] = {}
    for t in totals:
        k = _grouped_key(t.Category)
        if k not in grouped:
            grouped[k] = MonthlyEmployeeTotal(k, 0.0, 0.0, 0.0, 0.0, 0.0)
        g = grouped[k]
        g.BasicHours += t.BasicHours
        g.MonFriOvertime += t.MonFriOvertime
        g.SatSunOvertime += t.SatSunOvertime
        g.AnnualHoliday += t.AnnualHoliday
        g.TotalPaidHours += t.TotalPaidHours
    return grouped


def _enrich_week_summary(out: MonthlyWeekSummary) -> None:
    if not out.employee_totals and out.employees:
        out.employee_totals = _employee_totals_from_employees(out.employees)
    out.non_agency_total = sum(t.TotalPaidHours for t in out.employee_totals if not _is_agency_category(t.Category))
    if not out.grouped_totals:
        out.grouped_totals = _build_grouped_totals(out.employee_totals)
    out.emp_agency_bands = _compute_emp_agency_bands(out.employees)


def _workbook_has_all_data(file_obj: Any) -> bool:
    file_obj.seek(0)
    try:
        xl = pd.ExcelFile(file_obj)
        return "All Data" in xl.sheet_names
    except Exception:
        return False
    finally:
        file_obj.seek(0)


def parse_weekly_gazebo_all_data(
    file_obj: Any,
    *,
    start_date: str = "",
    end_date: str = "",
) -> MonthlyWeekSummary:
    """Parse weekly Gazebo export (.xlsx) — employees from the All Data sheet."""
    file_obj.seek(0)
    df = pd.read_excel(file_obj, sheet_name="All Data", header=None, dtype=str)
    text_rows = [[_to_text(v) for v in row] for row in df.values.tolist()]
    out = MonthlyWeekSummary(start_date=start_date, end_date=end_date)
    if not text_rows:
        return out

    header_row = -1
    for i, row in enumerate(text_rows[:5]):
        if _to_text(row[0] if len(row) > 0 else "").lower() == "name" and _to_text(row[1] if len(row) > 1 else "").lower() == "category":
            header_row = i
            break
    if header_row < 0:
        return out

    for row in text_rows[header_row + 1 :]:
        name = _to_text(row[0] if len(row) > 0 else "")
        if not name:
            break
        col_b = _to_text(row[1] if len(row) > 1 else "")
        if col_b.startswith("Category breakdown"):
            break
        out.employees.append(
            MonthlyEmployee(
                Name=name,
                Category=col_b,
                SageNo=_parse_int(row[2] if len(row) > 2 else ""),
                BasicHours=_parse_decimal(row[3] if len(row) > 3 else ""),
                MonFriOvertime=_parse_decimal(row[4] if len(row) > 4 else ""),
                SatSunOvertime=_parse_decimal(row[5] if len(row) > 5 else ""),
                AnnualHoliday=_parse_decimal(row[6] if len(row) > 6 else ""),
                TotalPaidHours=_parse_decimal(row[7] if len(row) > 7 else ""),
            )
        )
    _enrich_week_summary(out)
    return out


def parse_monthly_week_file(file_obj: Any) -> MonthlyWeekSummary:
    from .payroll_service import _load_sheet  # reuse weekly reader

    table = _load_sheet(file_obj)
    rows = table.values.tolist()
    text_rows = [[_to_text(v) for v in row] for row in rows]
    if not text_rows:
        return MonthlyWeekSummary()

    out = MonthlyWeekSummary()
    out.start_date = _to_text(text_rows[0][4] if len(text_rows[0]) > 4 else "").removeprefix("D ").strip()
    out.end_date = _to_text(text_rows[0][7] if len(text_rows[0]) > 7 else "").removeprefix("D ").strip()

    r = 3  # row 4 in excel
    while r < len(text_rows):
        row = text_rows[r]
        if not _to_text(row[0] if len(row) > 0 else ""):
            break
        out.employees.append(
            MonthlyEmployee(
                Name=_to_text(row[0] if len(row) > 0 else ""),
                Category=_to_text(row[1] if len(row) > 1 else ""),
                SageNo=_parse_int(row[2] if len(row) > 2 else ""),
                BasicHours=_parse_decimal(row[3] if len(row) > 3 else ""),
                MonFriOvertime=_parse_decimal(row[4] if len(row) > 4 else ""),
                SatSunOvertime=_parse_decimal(row[5] if len(row) > 5 else ""),
                AnnualHoliday=_parse_decimal(row[6] if len(row) > 6 else ""),
                TotalPaidHours=_parse_decimal(row[7] if len(row) > 7 else ""),
            )
        )
        r += 1

    adjustments_header = -1
    for i in range(r, min(len(text_rows), r + 120)):
        if _to_text(text_rows[i][1] if len(text_rows[i]) > 1 else "").startswith("Adjustments"):
            adjustments_header = i
            break
    if adjustments_header >= 0:
        ar = adjustments_header + 2
        while ar < len(text_rows):
            name = _to_text(text_rows[ar][1] if len(text_rows[ar]) > 1 else "")
            if not name:
                break
            out.adjustments.append(
                MonthlyAdjustment(
                    Name=name,
                    Type=_to_text(text_rows[ar][2] if len(text_rows[ar]) > 2 else ""),
                    Value=_parse_decimal(text_rows[ar][3] if len(text_rows[ar]) > 3 else ""),
                )
            )
            ar += 1
        r = ar

    totals_header = -1
    for i in range(r, min(len(text_rows), r + 150)):
        if _to_text(text_rows[i][1] if len(text_rows[i]) > 1 else "") == "Category":
            totals_header = i
            break
    if totals_header >= 0:
        tr = totals_header + 1
        while tr < len(text_rows):
            cat = _to_text(text_rows[tr][1] if len(text_rows[tr]) > 1 else "")
            if not cat:
                break
            total = MonthlyEmployeeTotal(
                Category=cat,
                BasicHours=_parse_decimal(text_rows[tr][3] if len(text_rows[tr]) > 3 else ""),
                MonFriOvertime=_parse_decimal(text_rows[tr][4] if len(text_rows[tr]) > 4 else ""),
                SatSunOvertime=_parse_decimal(text_rows[tr][5] if len(text_rows[tr]) > 5 else ""),
                AnnualHoliday=_parse_decimal(text_rows[tr][6] if len(text_rows[tr]) > 6 else ""),
                TotalPaidHours=_parse_decimal(text_rows[tr][7] if len(text_rows[tr]) > 7 else ""),
            )
            out.employee_totals.append(total)
            tr += 1

    _enrich_week_summary(out)
    return out


def parse_monthly_inputs(
    weekly_files: list[Any],
    week_dates: list[tuple[str, str]] | None = None,
) -> list[MonthlyWeekSummary]:
    week_dates = week_dates or []
    summaries: list[MonthlyWeekSummary] = []
    for i, f in enumerate(weekly_files):
        start_date, end_date = week_dates[i] if i < len(week_dates) else ("", "")
        f.seek(0)
        if _workbook_has_all_data(f):
            f.seek(0)
            s = parse_weekly_gazebo_all_data(f, start_date=start_date, end_date=end_date)
        else:
            f.seek(0)
            s = parse_monthly_week_file(f)
            if start_date:
                s.start_date = start_date
            if end_date:
                s.end_date = end_date
        summaries.append(s)
    return summaries


def monthly_summaries_to_json(summaries: list[MonthlyWeekSummary]) -> list[dict[str, Any]]:
    return [asdict(s) for s in summaries]


def monthly_summaries_from_json(data: list[dict[str, Any]]) -> list[MonthlyWeekSummary]:
    out: list[MonthlyWeekSummary] = []
    for d in data:
        s = MonthlyWeekSummary(
            employees=[MonthlyEmployee(**e) for e in d.get("employees", [])],
            employee_totals=[MonthlyEmployeeTotal(**t) for t in d.get("employee_totals", [])],
            adjustments=[MonthlyAdjustment(**a) for a in d.get("adjustments", [])],
            start_date=str(d.get("start_date", "")),
            end_date=str(d.get("end_date", "")),
            non_agency_total=float(d.get("non_agency_total", 0.0)),
            emp_agency_bands=dict(d.get("emp_agency_bands") or {}),
        )
        s.grouped_totals = {}
        for k, v in (d.get("grouped_totals") or {}).items():
            if isinstance(v, dict):
                s.grouped_totals[str(k)] = MonthlyEmployeeTotal(**v)
        if not s.emp_agency_bands and s.employees:
            _enrich_week_summary(s)
        out.append(s)
    return out


def _write_header(ws, r: int, c: int = 1) -> None:
    ws.cell(r, c, "Name")
    ws.cell(r, c + 1, "Category")
    ws.cell(r, c + 2, "SageNo")
    ws.cell(r, c + 3, "BasicHours")
    ws.cell(r, c + 4, "MonFriOvertime")
    ws.cell(r, c + 5, "SatSunOvertime")
    ws.cell(r, c + 6, "AnnualHoliday")
    ws.cell(r, c + 7, "TotalPaidHours")


def _write_total_header(ws, r: int) -> None:
    ws.cell(r, 2, "Category")
    ws.cell(r, 4, "BasicHours")
    ws.cell(r, 5, "MonFriOvertime")
    ws.cell(r, 6, "SatSunOvertime")
    ws.cell(r, 7, "AnnualHoliday")
    ws.cell(r, 8, "TotalPaidHours")


def _write_emp_agency_block(
    ws,
    r: int,
    bands: dict[str, dict[str, float]],
    *,
    row_label: str | None = None,
) -> int:
    """EMP/AGENCY/TOTAL in col C with bands in D–H; optional label in col B on first row."""
    for i, key in enumerate(_EMP_AGENCY_ROWS):
        if row_label and i == 0:
            ws.cell(r, 2, row_label)
        ws.cell(r, 3, key)
        row_bands = bands.get(key) or _empty_bands()
        for j, col in enumerate(_BAND_KEYS):
            ws.cell(r, 4 + j, float(row_bands.get(col, 0.0) or 0.0))
        r += 1
    return r


def build_monthly_excel_bytes(
    week_summaries: list[MonthlyWeekSummary],
    non_hourly_names: set[str] | None = None,
) -> bytes:
    wb = Workbook()
    wb.remove(wb.active)

    grouped_from_weeks: list[dict[str, MonthlyEmployeeTotal]] = []
    week_emp_agency_list: list[dict[str, dict[str, float]]] = []

    for i, s in enumerate(week_summaries, start=1):
        if not s.emp_agency_bands:
            _enrich_week_summary(s)
        week_emp_agency_list.append(s.emp_agency_bands)

        ws = wb.create_sheet(f"Week{i}")
        ws.cell(1, 1, f"Week {i}")
        ws.cell(1, 4, "Start Date")
        ws.cell(1, 5, f"D {s.start_date}" if s.start_date else "")
        ws.cell(1, 7, "End Date")
        ws.cell(1, 8, f"D {s.end_date}" if s.end_date else "")
        _write_header(ws, 3)

        r = 4
        for e in s.employees:
            ws.cell(r, 1, e.Name)
            ws.cell(r, 2, e.Category)
            ws.cell(r, 3, e.SageNo)
            ws.cell(r, 4, e.BasicHours)
            ws.cell(r, 5, e.MonFriOvertime)
            ws.cell(r, 6, e.SatSunOvertime)
            ws.cell(r, 7, e.AnnualHoliday)
            ws.cell(r, 8, e.TotalPaidHours)
            r += 1

        r += 2
        ws.cell(r, 2, f"Adjustments - {len(s.adjustments)}")
        r += 1
        ws.cell(r, 2, "Name")
        ws.cell(r, 3, "Type")
        ws.cell(r, 4, "Value")
        r += 1
        for a in s.adjustments:
            ws.cell(r, 2, a.Name)
            ws.cell(r, 3, a.Type)
            ws.cell(r, 4, a.Value)
            r += 1

        r += 2
        _write_total_header(ws, r)
        r += 1
        for t in s.employee_totals:
            ws.cell(r, 2, t.Category)
            ws.cell(r, 4, t.BasicHours)
            ws.cell(r, 5, t.MonFriOvertime)
            ws.cell(r, 6, t.SatSunOvertime)
            ws.cell(r, 7, t.AnnualHoliday)
            ws.cell(r, 8, t.TotalPaidHours)
            r += 1

        r += 1
        for k in sorted(s.grouped_totals.keys()):
            v = s.grouped_totals[k]
            ws.cell(r, 2, k)
            ws.cell(r, 4, v.BasicHours)
            ws.cell(r, 5, v.MonFriOvertime)
            ws.cell(r, 6, v.SatSunOvertime)
            ws.cell(r, 7, v.AnnualHoliday)
            ws.cell(r, 8, v.TotalPaidHours)
            r += 1

        r += 2
        r = _write_emp_agency_block(ws, r, s.emp_agency_bands)
        grouped_from_weeks.append(s.grouped_totals)

    merged_employees: dict[str, MonthlyEmployee] = {}
    merged_totals: dict[str, MonthlyEmployeeTotal] = {}
    for s in week_summaries:
        for e in s.employees:
            cur = merged_employees.get(e.Name)
            if cur is None:
                merged_employees[e.Name] = MonthlyEmployee(**e.__dict__)
                continue
            cur.BasicHours += e.BasicHours
            cur.MonFriOvertime += e.MonFriOvertime
            cur.SatSunOvertime += e.SatSunOvertime
            cur.AnnualHoliday += e.AnnualHoliday
            cur.TotalPaidHours += e.TotalPaidHours
            cur.SageNo = e.SageNo
            cur.Category = e.Category

        for t in s.employee_totals:
            cur_t = merged_totals.get(t.Category)
            if cur_t is None:
                merged_totals[t.Category] = MonthlyEmployeeTotal(**t.__dict__)
                continue
            cur_t.BasicHours += t.BasicHours
            cur_t.MonFriOvertime += t.MonFriOvertime
            cur_t.SatSunOvertime += t.SatSunOvertime
            cur_t.AnnualHoliday += t.AnnualHoliday
            cur_t.TotalPaidHours += t.TotalPaidHours

    non_hourly_names = {n.strip().upper() for n in (non_hourly_names or set()) if n and n.strip()}
    for e in merged_employees.values():
        if e.Name.strip().upper() in non_hourly_names:
            e.IsHourly = False
        if not e.IsHourly and e.Category.strip().startswith("A-"):
            raise ValueError(f"Agency employee cannot be non-hourly: {e.Name}")

    ws = wb.create_sheet("Summary")
    ws.cell(1, 1, "Week Monthly Summary")
    if week_summaries and week_summaries[0].start_date:
        ws.cell(1, 4, "Start Date")
        ws.cell(1, 5, f"D {week_summaries[0].start_date}")
    if week_summaries and week_summaries[-1].end_date:
        ws.cell(1, 7, "End Date")
        ws.cell(1, 8, f"D {week_summaries[-1].end_date}")

    _write_header(ws, 3)
    r = 4
    for e in merged_employees.values():
        ws.cell(r, 1, e.Name)
        ws.cell(r, 2, e.Category)
        ws.cell(r, 3, e.SageNo)
        ws.cell(r, 4, e.BasicHours)
        ws.cell(r, 5, e.MonFriOvertime)
        ws.cell(r, 6, e.SatSunOvertime)
        ws.cell(r, 7, e.AnnualHoliday)
        ws.cell(r, 8, e.TotalPaidHours)
        r += 1

    r += 2
    _write_total_header(ws, r)
    r += 1
    for t in merged_totals.values():
        ws.cell(r, 2, t.Category)
        ws.cell(r, 4, t.BasicHours)
        ws.cell(r, 5, t.MonFriOvertime)
        ws.cell(r, 6, t.SatSunOvertime)
        ws.cell(r, 7, t.AnnualHoliday)
        ws.cell(r, 8, t.TotalPaidHours)
        non_hourly = [e for e in merged_employees.values() if not e.IsHourly and e.Category.upper() == t.Category.upper()]
        if non_hourly:
            ws.cell(r + 1, 2, f"{t.Category} non-hourly hours")
            ws.cell(r + 1, 4, -sum(e.BasicHours for e in non_hourly))
            ws.cell(r + 1, 5, -sum(e.MonFriOvertime for e in non_hourly))
            ws.cell(r + 1, 6, -sum(e.SatSunOvertime for e in non_hourly))
            ws.cell(r + 1, 7, -sum(e.AnnualHoliday for e in non_hourly))
            ws.cell(r + 1, 8, -sum(e.TotalPaidHours for e in non_hourly))
            r += 2
        else:
            r += 1

    r += 1
    agg_grouped: dict[str, MonthlyEmployeeTotal] = {}
    for m in grouped_from_weeks:
        for k, t in m.items():
            if k not in agg_grouped:
                agg_grouped[k] = MonthlyEmployeeTotal(k, 0.0, 0.0, 0.0, 0.0, 0.0)
            a = agg_grouped[k]
            a.BasicHours += t.BasicHours
            a.MonFriOvertime += t.MonFriOvertime
            a.SatSunOvertime += t.SatSunOvertime
            a.AnnualHoliday += t.AnnualHoliday
            a.TotalPaidHours += t.TotalPaidHours
    for k in sorted(agg_grouped.keys()):
        t = agg_grouped[k]
        ws.cell(r, 2, t.Category)
        ws.cell(r, 4, t.BasicHours)
        ws.cell(r, 5, t.MonFriOvertime)
        ws.cell(r, 6, t.SatSunOvertime)
        ws.cell(r, 7, t.AnnualHoliday)
        ws.cell(r, 8, t.TotalPaidHours)
        r += 1

    if week_emp_agency_list:
        r += 2
        monthly_bands = _sum_emp_agency_bands(week_emp_agency_list)
        r = _write_emp_agency_block(ws, r, monthly_bands, row_label="MONTHLY")

        for wi, bands in enumerate(week_emp_agency_list):
            r += 1
            label = "Weekly" if wi == 0 else None
            r = _write_emp_agency_block(ws, r, bands, row_label=label)

        merged_month = _compute_emp_agency_bands(list(merged_employees.values()))
        week_total_sum = _empty_bands()
        for bands in week_emp_agency_list:
            for k in _BAND_KEYS:
                week_total_sum[k] += float(bands.get("TOTAL", {}).get(k, 0.0) or 0.0)
        diff_bands = {
            "EMP": {k: merged_month["EMP"][k] - monthly_bands["EMP"][k] for k in _BAND_KEYS},
            "AGENCY": {k: merged_month["AGENCY"][k] - monthly_bands["AGENCY"][k] for k in _BAND_KEYS},
            "TOTAL": {k: merged_month["TOTAL"][k] - week_total_sum[k] for k in _BAND_KEYS},
        }
        r += 1
        r = _write_emp_agency_block(ws, r, diff_bands, row_label="Diff")

    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return out.getvalue()
