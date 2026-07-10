"""Backward-compatible facade for weekly payroll APIs.

Callers should keep importing from this module; implementations live under
parsing/, payroll/, and excel/.
"""

from __future__ import annotations

from .excel.workbook import (
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
from .parsing.common import _load_sheet, _normalize_header
from .parsing.contract import (
    ContractAuditResult,
    _resolve_contracted_hours,
    audit_contract_integrity,
    audit_contract_pay_id_coverage,
    load_contract_file_index,
    parse_contracted_hours,
    parse_employee_display_names,
)
from .parsing.employee_hours import parse_employee_hours, parse_processing_date
from .payroll.calculate import (
    HOLIDAY_PAY_FACTOR,
    calculate_payroll,
    calculate_weekly_payroll,
    compute_extra_holiday_pay,
)
from .payroll.summary import (
    agency_categories_from_rows,
    build_staff_summary,
    is_agency_category,
    split_emp_agency_rows,
    total_paid_hours_from_rows,
)
from .payroll.types import PayrollResult

__all__ = [
    "ContractAuditResult",
    "HOLIDAY_PAY_FACTOR",
    "PayrollResult",
    "_HOUR_BAND_COLS",
    "_OVERALL_CATEGORY_ORDER",
    "_build_analysis_dataframe",
    "_build_grouped_analysis_dataframe",
    "_build_overall_analysis_dataframe",
    "_load_sheet",
    "_normalize_header",
    "_overall_category_key",
    "_resolve_contracted_hours",
    "_sum_hour_bands",
    "agency_categories_from_rows",
    "audit_contract_integrity",
    "audit_contract_pay_id_coverage",
    "build_category_summary_hr_df",
    "build_emp_agency_total_df",
    "build_excel_bytes",
    "build_hours_over_60_df",
    "build_overall_category_totals",
    "build_staff_summary",
    "calculate_payroll",
    "calculate_weekly_payroll",
    "compute_extra_holiday_pay",
    "is_agency_category",
    "load_contract_file_index",
    "parse_contracted_hours",
    "parse_employee_display_names",
    "parse_employee_hours",
    "parse_processing_date",
    "split_emp_agency_rows",
    "total_paid_hours_from_rows",
]
