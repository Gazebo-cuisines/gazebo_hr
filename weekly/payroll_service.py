from __future__ import annotations

import logging
from dataclasses import dataclass
from io import BytesIO
from typing import Any

import pandas as pd
from openpyxl.styles import Border, Side


logger = logging.getLogger(__name__)

_TABLE_BORDER = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)


# ClockRite "Paid Hours (Inc Absence) Summary": Pay ID column B (0-based 1); annual
# hours duplicated at Excel H and L (0-based 7 and 11). See payroll fix plan.
_CLOCKRITE_PAY_ID_COL = 1
_CLOCKRITE_SAGE_HEADER_COL = 3
_CLOCKRITE_ANNUAL_H_COL = 7
_CLOCKRITE_ANNUAL_L_COL = 11

# Appended category breakdown / grand total (matches monthly export layout).
_CATEGORY_COL = 2
_CATEGORY_NUM_START = 4
_CATEGORY_BAND_COLS = ["BasicHours", "MonFriOvertime", "SatSunOvertime", "AnnualHoliday", "TotalPaidHours"]

AGENCY_CATEGORIES = {
    "A-EL CLNR",
    "A-EL DPCH",
    "A-EL PKNG HIGH_RISK",
    "A-EL PKNG SLEEVING",
    "A-EL PROD BELT",
    "A-EL PROD FORMING",
    "A-EL PROD FRYING",
    "A-EL PROD PREPARATION",
    "A-EL TECHNICAL",
    "A-PM PKNG HIGH_RISK",
    "A-RS PKNG SLEEVING",
    "A-RS PROD BELT",
    "A-RS PROD FRYING",
}


@dataclass
class PayrollResult:
    rows: list[dict[str, Any]]
    agency_rows: list[dict[str, Any]]
    gazebo_rows: list[dict[str, Any]]
    total_paid_hours: float


def total_paid_hours_from_rows(rows: list[dict[str, Any]]) -> float:
    """Sum TotalPaidHours across all employee rows (includes every category, including A- prefix)."""
    return round(sum(float(r.get("TotalPaidHours", 0.0)) for r in rows), 2)


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

        result.append(
            {
                "Name": name_text.upper(),
                "Category": category,
                "SageNo": pay_id,
                "BasicHours": basic_hours,
                "MonFriOvertime": mon_fri_ot,
                "SatSunOvertime": sat_sun_ot,
                "AnnualHoliday": annual_holiday,
                "TotalPaidHours": total_paid,
            }
        )

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


def _find_header_index(headers: list[str], *candidates: str) -> int:
    for idx, cell in enumerate(headers):
        n = _normalize_header(cell)
        if any(n == c for c in candidates):
            return idx
    return -1


def _contract_hrs_value_in_row(row: list[str]) -> float | None:
    for c, cell in enumerate(row):
        n = _normalize_header(cell)
        if n in ("contracthrs", "contracthours", "contracthr", "contractedhours"):
            if c + 1 < len(row):
                return _parse_decimal(row[c + 1])
            return 0.0
    return None


def _sage_pay_ref_ids_near(rows: list[list[str]], r_payroll: int, row_span: int = 22, row_forward: int = 8) -> list[int]:
    """ClockRite employee blocks may put Sage Pay Ref on another row than Payroll Number (Sage ≠ Prox ID)."""
    seen: set[int] = set()
    out: list[int] = []
    r0 = max(0, r_payroll - row_span)
    r1 = min(len(rows), r_payroll + row_forward)
    for rr in range(r0, r1):
        row = rows[rr]
        for c, cell in enumerate(row):
            if _normalize_header(cell) != "sagepayref":
                continue
            if c + 1 >= len(row):
                continue
            sid = _parse_int(row[c + 1])
            if sid is None or sid <= 0:
                continue
            if sid not in seen:
                seen.add(sid)
                out.append(sid)
    return out


def _block_start_row(rows: list[list[str]], r_payroll: int) -> int:
    """Row index of the block title row (Prox ID + full name) above Payroll Number."""
    for back in range(1, 30):
        if r_payroll - back < 0:
            break
        pr = rows[r_payroll - back]
        if not pr:
            continue
        a0 = _to_text(pr[0])
        if a0.isdigit() and len(a0) < 6 and len(pr) > 1 and _to_text(pr[1]):
            return r_payroll - back
    return max(0, r_payroll - 25)


def _sage_pay_ref_ids_in_block(rows: list[list[str]], r_payroll: int) -> list[int]:
    """Sage Pay Ref values within one employee block only (avoids adjacent-block bleed)."""
    seen: set[int] = set()
    out: list[int] = []
    r0 = _block_start_row(rows, r_payroll)
    r1 = min(len(rows), r_payroll + 4)
    for rr in range(r0, r1):
        row = rows[rr]
        for c, cell in enumerate(row):
            if _normalize_header(cell) != "sagepayref":
                continue
            if c + 1 >= len(row):
                continue
            sid = _parse_int(row[c + 1])
            if sid is None or sid <= 0:
                continue
            if sid not in seen:
                seen.add(sid)
                out.append(sid)
    return out


def _block_names_near(rows: list[list[str]], r_payroll: int) -> tuple[str, str]:
    """Full name from block title row; clock name from Clock Name row above Payroll Number."""
    full_name = ""
    clock_name = ""
    for back in range(1, 30):
        if r_payroll - back < 0:
            break
        pr = rows[r_payroll - back]
        if not pr:
            continue
        for c, cell in enumerate(pr):
            if _normalize_header(cell) == "clockname" and c + 1 < len(pr):
                cn = _to_text(pr[c + 1])
                if cn:
                    clock_name = cn
        if full_name:
            continue
        a0 = _to_text(pr[0])
        if a0.isdigit() and len(a0) < 6 and len(pr) > 1:
            candidate = _to_text(pr[1])
            if candidate:
                full_name = candidate
    return full_name, clock_name


def _parse_clockrite_display_names(rows: list[list[str]]) -> tuple[dict[int, str], dict[str, str]]:
    """Map Sage/Payroll IDs and clock names to full names from Employee Details blocks."""
    by_sage: dict[int, str] = {}
    by_clock: dict[str, str] = {}
    for r, row in enumerate(rows):
        for c, cell in enumerate(row):
            if _normalize_header(cell) != "payrollnumber":
                continue
            if c + 1 >= len(row):
                continue
            pay_no = _parse_int(row[c + 1])
            if pay_no is None or pay_no == 0:
                continue
            full_name, clock_name = _block_names_near(rows, r)
            if not full_name:
                continue
            by_sage[pay_no] = full_name
            for sid in _sage_pay_ref_ids_in_block(rows, r):
                by_sage[sid] = full_name
            if clock_name:
                by_clock[clock_name.upper()] = full_name
    return by_sage, by_clock


def parse_employee_display_names(file_obj: Any) -> tuple[dict[int, str], dict[str, str]]:
    """Full names from contract file (ClockRite Employee Details layout)."""
    df = _load_sheet(file_obj)
    rows = [[_to_text(v) for v in rec] for rec in df.values.tolist()]
    if not rows:
        return {}, {}
    return _parse_clockrite_display_names(rows)


def _parse_clockrite_contract_report(rows: list[list[str]]) -> tuple[dict[int, float], dict[str, float]]:
    """
    Clockrite "Employee Details - Advanced" XLS: repeated blocks with label rows
    (e.g. "Contract Hrs" and "Payroll Number" in separate rows, not one header table).
    Join key: Payroll Number value == employee file Pay ID (SageNo).
    """
    by_payroll: dict[int, float] = {}
    by_name: dict[str, float] = {}
    for r, row in enumerate(rows):
        for c, cell in enumerate(row):
            if _normalize_header(cell) != "payrollnumber":
                continue
            if c + 1 >= len(row):
                continue
            pay_no = _parse_int(row[c + 1])
            if pay_no is None or pay_no == 0:
                continue
            hours = 0.0
            for back in range(1, 25):
                if r - back < 0:
                    break
                prev = rows[r - back]
                got = _contract_hrs_value_in_row(prev)
                if got is not None:
                    hours = got
                    break
            by_payroll[pay_no] = hours
            for back in range(1, 30):
                if r - back < 0:
                    break
                pr = rows[r - back]
                if not pr:
                    continue
                a0 = _to_text(pr[0])
                if a0.isdigit() and len(a0) < 6 and len(pr) > 1:
                    name_key = _to_text(pr[1]).upper()
                    if name_key:
                        by_name[name_key] = hours
                    break
            for sid in _sage_pay_ref_ids_near(rows, r):
                if sid == pay_no:
                    continue
                by_payroll[sid] = hours
                logger.debug(
                    "Contract hours: SagePayRef %s aliased to PayrollNumber %s hours=%s",
                    sid,
                    pay_no,
                    hours,
                )
    return by_payroll, by_name


def parse_contracted_hours(file_obj: Any) -> tuple[dict[int, float], dict[str, float]]:
    df = _load_sheet(file_obj)
    rows = [[_to_text(v) for v in rec] for rec in df.values.tolist()]
    if not rows:
        return {}, {}

    header_row = -1
    payroll_col = -1
    contract_col = -1
    name_col = -1
    for r, row in enumerate(rows[:80]):
        pc = _find_header_index(row, "payrollnumber", "payrollno", "payroll")
        cc = _find_header_index(row, "contracthrs", "contracthours", "contracthr", "contractedhours")
        if pc >= 0 and cc >= 0:
            header_row = r
            payroll_col = pc
            contract_col = cc
            name_col = _find_header_index(row, "clockname", "name", "employeename")
            break

    if header_row >= 0:
        by_payroll: dict[int, float] = {}
        by_name: dict[str, float] = {}
        for row in rows[header_row + 1 :]:
            if _row_contains_date_range(row):
                break
            hours = _parse_decimal(row[contract_col] if contract_col < len(row) else "")
            if hours == 0.0:
                continue
            pay_no = _parse_int(row[payroll_col] if payroll_col < len(row) else "")
            if pay_no is not None:
                by_payroll[pay_no] = hours
            if name_col >= 0 and name_col < len(row):
                name = _to_text(row[name_col]).upper()
                if name:
                    by_name[name] = hours
        if by_payroll or by_name:
            return by_payroll, by_name

    return _parse_clockrite_contract_report(rows)


def audit_contract_pay_id_coverage(
    employee_rows: list[dict[str, Any]], contracted_file_obj: Any
) -> list[dict[str, Any]]:
    """Employees whose Pay ID does not appear in the contract export (Payroll Number / Sage Pay Ref)."""
    contracted_file_obj.seek(0)
    by_payroll, _ = parse_contracted_hours(contracted_file_obj)
    missing: list[dict[str, Any]] = []
    for row in employee_rows:
        sage_no = int(row["SageNo"])
        if sage_no not in by_payroll:
            missing.append(
                {
                    "SageNo": sage_no,
                    "Name": row.get("Name"),
                    "Category": row.get("Category"),
                }
            )
    return missing


def _resolve_contracted_hours(
    sage_no: int,
    name_upper: str,
    by_payroll: dict[int, float],
    by_name: dict[str, float],
    by_clock_name: dict[str, str],
) -> tuple[float, str, str]:
    """Return (contracted hours, ContractHourMatch, ContractMatchReason)."""
    hit = by_payroll.get(sage_no)
    if hit is not None:
        return float(hit), "Yes", "Matched on Pay ID"
    hit = by_name.get(name_upper)
    if hit:
        return float(hit), "Yes", "Matched on employee name"
    full = by_clock_name.get(name_upper)
    if full:
        hit = by_name.get(full.upper())
        if hit:
            return float(hit), "Yes", "Matched on employee name (clock name lookup)"
    return 0.0, "No", "Pay ID not in contract export"


def calculate_payroll(employee_rows: list[dict[str, Any]], contracted_file_obj: Any) -> PayrollResult:
    contracted_file_obj.seek(0)
    by_payroll, by_name = parse_contracted_hours(contracted_file_obj)
    contracted_file_obj.seek(0)
    by_sage_name, by_clock_name = parse_employee_display_names(contracted_file_obj)

    for row in employee_rows:
        name_upper = str(row["Name"]).upper()
        contracted, match, reason = _resolve_contracted_hours(
            int(row["SageNo"]), name_upper, by_payroll, by_name, by_clock_name
        )
        row["ContractHourMatch"] = match
        row["ContractMatchReason"] = reason
        if match == "No":
            logger.debug(
                "Contract match miss: SageNo=%s Name=%s (%s)",
                row.get("SageNo"),
                row.get("Name"),
                reason,
            )
        row["ContractedHours"] = contracted
        row["Overtime"] = max(0.0, float(row["TotalPaidHours"]) - contracted)

    for row in employee_rows:
        full = by_sage_name.get(int(row["SageNo"])) or by_clock_name.get(str(row["Name"]).upper())
        if full:
            row["Name"] = full

    agency_rows = [r for r in employee_rows if str(r["Category"]).strip().upper() in AGENCY_CATEGORIES]
    gazebo_rows = [r for r in employee_rows if str(r["Category"]).strip().upper() not in AGENCY_CATEGORIES]
    all_hours = total_paid_hours_from_rows(employee_rows)

    return PayrollResult(
        rows=employee_rows,
        agency_rows=agency_rows,
        gazebo_rows=gazebo_rows,
        total_paid_hours=all_hours,
    )


_HOUR_BAND_COLS = ["BasicHours", "MonFriOvertime", "SatSunOvertime", "AnnualHoliday", "TotalPaidHours"]


def _sum_hour_bands(rows: list[dict[str, Any]]) -> dict[str, float]:
    return {k: sum(float(r.get(k, 0.0) or 0.0) for r in rows) for k in _HOUR_BAND_COLS}


def build_emp_agency_total_df(result: PayrollResult) -> pd.DataFrame:
    """EMP (non-agency / Gazebo), AGENCY, and TOTAL sums for the five hour bands (section C)."""
    emp = _sum_hour_bands(result.gazebo_rows)
    agency = _sum_hour_bands(result.agency_rows)
    total = _sum_hour_bands(result.rows)
    return pd.DataFrame(
        [
            {"Summary": "EMP", **emp},
            {"Summary": "AGENCY", **agency},
            {"Summary": "TOTAL", **total},
        ]
    )


def build_category_summary_hr_df(analysis_df: pd.DataFrame) -> pd.DataFrame:
    """Same data as Analysis with HR-friendly column names and a grand total row (section B)."""
    hr_cols = {
        "BasicHours": "Basic Hour",
        "MonFriOvertime": "Mon Fri O",
        "SatSunOvertime": "Sat Sun O",
        "AnnualHoliday": "Annual Ho",
        "TotalPaidHours": "Total Paid Hours",
    }
    if analysis_df.empty:
        return pd.DataFrame(columns=["Category", *list(hr_cols.values())])
    out = analysis_df.rename(columns=hr_cols)
    total_row: dict[str, Any] = {"Category": "Grand total"}
    for src, dst in hr_cols.items():
        total_row[dst] = float(analysis_df[src].sum())
    return pd.concat([out, pd.DataFrame([total_row])], ignore_index=True)


def build_hours_over_60_df(all_df: pd.DataFrame) -> pd.DataFrame:
    """Employees with TotalPaidHours strictly greater than 60."""
    if all_df.empty:
        return pd.DataFrame(columns=["Name", "Category", "SageNo", "TotalPaidHours", "BasicHours", "Overtime", "ContractedHours"])
    work = all_df.copy()
    work["_tp"] = pd.to_numeric(work["TotalPaidHours"], errors="coerce").fillna(0.0)
    filtered = work[work["_tp"] > 60.0].drop(columns=["_tp"], errors="ignore")
    cols = ["Name", "Category", "SageNo", "TotalPaidHours", "BasicHours", "Overtime", "ContractedHours"]
    have = [c for c in cols if c in filtered.columns]
    return filtered[have] if have else filtered


def _build_analysis_dataframe(all_df: pd.DataFrame) -> pd.DataFrame:
    if all_df.empty:
        return pd.DataFrame(columns=["Category", "BasicHours", "MonFriOvertime", "SatSunOvertime", "AnnualHoliday", "TotalPaidHours"])
    return (
        all_df.groupby("Category", dropna=False)[["BasicHours", "MonFriOvertime", "SatSunOvertime", "AnnualHoliday", "TotalPaidHours"]]
        .sum()
        .reset_index()
    )


def _apply_table_border(ws, min_row: int, max_row: int, min_col: int, max_col: int) -> None:
    """Apply thin borders to every cell in the given range (formatting only)."""
    for r in range(min_row, max_row + 1):
        for c in range(min_col, max_col + 1):
            ws.cell(r, c).border = _TABLE_BORDER


def _append_grand_total_row_openpyxl(ws, analysis_df: pd.DataFrame, start_row: int) -> int:
    """Write 'Grand total' and column sums. Returns one past the last row written."""
    if analysis_df.empty:
        return start_row
    r = start_row
    ws.cell(r, _CATEGORY_COL, "Grand total")
    for i, col in enumerate(_CATEGORY_BAND_COLS):
        ws.cell(r, _CATEGORY_NUM_START + i, float(analysis_df[col].sum()))
    return r + 1


def _append_category_breakdown_block(ws, analysis_df: pd.DataFrame, start_row: int) -> int:
    """Write analysis headers, category rows, blank row, and grand total. Returns one past the last row."""
    if analysis_df.empty:
        return start_row
    r = start_row
    ws.cell(r, _CATEGORY_COL, "Category")
    for i, h in enumerate(_CATEGORY_BAND_COLS):
        ws.cell(r, _CATEGORY_NUM_START + i, h)
    r += 1
    for _, row in analysis_df.iterrows():
        cat = row["Category"]
        ws.cell(r, _CATEGORY_COL, "" if pd.isna(cat) else str(cat))
        for i, h in enumerate(_CATEGORY_BAND_COLS):
            v = row[h]
            ws.cell(r, _CATEGORY_NUM_START + i, 0.0 if pd.isna(v) else float(v))
        r += 1
    r += 1
    return _append_grand_total_row_openpyxl(ws, analysis_df, r)


def _append_emp_agency_total_block(ws, emp_agency_df: pd.DataFrame, start_row: int) -> int:
    """Write EMP / AGENCY / TOTAL summary below category block (cols B and D–H). Returns one past last row."""
    if emp_agency_df.empty:
        return start_row
    r = start_row
    ws.cell(r, _CATEGORY_COL, "Summary")
    for i, h in enumerate(_CATEGORY_BAND_COLS):
        ws.cell(r, _CATEGORY_NUM_START + i, h)
    r += 1
    for _, row in emp_agency_df.iterrows():
        summary = row["Summary"]
        ws.cell(r, _CATEGORY_COL, "" if pd.isna(summary) else str(summary))
        for i, h in enumerate(_CATEGORY_BAND_COLS):
            v = row[h]
            ws.cell(r, _CATEGORY_NUM_START + i, 0.0 if pd.isna(v) else float(v))
        r += 1
    return r


def build_excel_bytes(result: PayrollResult) -> bytes:
    all_df = pd.DataFrame(result.rows)
    agency_df = pd.DataFrame(result.agency_rows)
    gazebo_df = pd.DataFrame(result.gazebo_rows)
    analysis_df = _build_analysis_dataframe(all_df)
    emp_agency_df = build_emp_agency_total_df(result)
    category_hr_df = build_category_summary_hr_df(analysis_df)
    over60_df = build_hours_over_60_df(all_df)

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        all_df.to_excel(writer, sheet_name="All Data", index=False)
        agency_df.to_excel(writer, sheet_name="Agency Employee", index=False)
        gazebo_df.to_excel(writer, sheet_name="Gazebo Employee", index=False)
        analysis_df.to_excel(writer, sheet_name="Analysis", index=False)

        if not analysis_df.empty:
            ws_all = writer.book["All Data"]
            all_data_start = int(ws_all.max_row) + 2
            end_row = _append_category_breakdown_block(ws_all, analysis_df, all_data_start)
            _apply_table_border(
                ws_all,
                all_data_start,
                end_row - 1,
                _CATEGORY_COL,
                _CATEGORY_NUM_START + len(_CATEGORY_BAND_COLS) - 1,
            )
            _append_emp_agency_total_block(ws_all, emp_agency_df, end_row + 1)

            ws_an = writer.book["Analysis"]
            an_start = int(ws_an.max_row) + 2
            _append_grand_total_row_openpyxl(ws_an, analysis_df, an_start)

        emp_agency_df.to_excel(writer, sheet_name="EMP Agency Total", index=False)
        category_hr_df.to_excel(writer, sheet_name="Category summary", index=False)
        over60_df.to_excel(writer, sheet_name="Hours over 60", index=False)

    output.seek(0)
    return output.getvalue()
