from __future__ import annotations

from django.contrib import messages
from django.http import HttpRequest, HttpResponse
from django.shortcuts import redirect
from django.views.decorators.http import require_GET

from ..export_service import (
	EXPORT_COLUMNS,
	add_branding_cover_sheet,
	add_weekly_branding_cover_sheet,
	build_csv_bytes,
	build_pdf_bytes,
	build_weekly_csv_bytes,
	build_weekly_pdf_bytes,
	day_report_filename,
	month_report_date_label,
	month_report_filename,
	week_report_filename,
	WEEKLY_EXPORT_HEADER_LABELS,
)
from ..monthly_service import build_monthly_excel_bytes, monthly_summaries_from_json
from ..payroll_service import (
	PayrollResult,
	build_excel_bytes,
	build_staff_summary,
	split_emp_agency_rows,
	total_paid_hours_from_rows,
)


def _daily_session_data(request: HttpRequest) -> tuple[list[dict], dict]:
	result_data = request.session.get('daily_last_result', {})
	rows = result_data.get('rows', [])
	stored = result_data.get('summary', {})
	if not rows:
		return [], stored
	summary = build_staff_summary(rows)
	if stored.get('processing_date'):
		summary['processing_date'] = stored['processing_date']
	summary['operator'] = stored.get('operator', 'HR')
	return rows, summary


def _weekly_session_data(request: HttpRequest) -> tuple[list[dict], dict]:
	result_data = request.session.get('weekly_last_result', {})
	return result_data.get('rows', []), result_data.get('summary', {})


@require_GET
def download_daily_excel(request: HttpRequest):
	rows, summary = _daily_session_data(request)
	if not rows:
		messages.error(request, 'No processed data available. Upload files first.')
		return redirect('weekly:daily_report')

	gazebo_rows, agency_rows = split_emp_agency_rows(rows)
	payroll_result = PayrollResult(
		rows=rows,
		agency_rows=agency_rows,
		gazebo_rows=gazebo_rows,
		total_paid_hours=total_paid_hours_from_rows(rows),
	)
	file_bytes = build_excel_bytes(payroll_result, employee_columns=EXPORT_COLUMNS)
	try:
		file_bytes = add_branding_cover_sheet(file_bytes, summary=summary)
	except Exception:
		pass
	response = HttpResponse(
		file_bytes,
		content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
	)
	response['Content-Disposition'] = f'attachment; filename="{day_report_filename("xlsx", summary.get("processing_date"))}"'
	return response


@require_GET
def download_daily_csv(request: HttpRequest):
	rows, summary = _daily_session_data(request)
	if not rows:
		messages.error(request, 'No processed data available. Upload files first.')
		return redirect('weekly:daily_report')
	file_bytes = build_csv_bytes(rows, summary=summary)
	response = HttpResponse(file_bytes, content_type='text/csv; charset=utf-8')
	response['Content-Disposition'] = f'attachment; filename="{day_report_filename("csv", summary.get("processing_date"))}"'
	return response


@require_GET
def download_daily_pdf(request: HttpRequest):
	rows, summary = _daily_session_data(request)
	if not rows:
		messages.error(request, 'No processed data available. Upload files first.')
		return redirect('weekly:daily_report')
	file_bytes = build_pdf_bytes(rows, summary=summary)
	response = HttpResponse(file_bytes, content_type='application/pdf')
	response['Content-Disposition'] = f'attachment; filename="{day_report_filename("pdf", summary.get("processing_date"))}"'
	return response


@require_GET
def download_weekly_excel(request: HttpRequest):
	rows, summary = _weekly_session_data(request)
	if not rows:
		messages.error(request, 'No processed data available. Upload files first.')
		return redirect('weekly:weekly_report')

	gazebo_rows, agency_rows = split_emp_agency_rows(rows)
	payroll_result = PayrollResult(
		rows=rows,
		agency_rows=agency_rows,
		gazebo_rows=gazebo_rows,
		total_paid_hours=total_paid_hours_from_rows(rows),
	)
	file_bytes = build_excel_bytes(payroll_result, column_rename=WEEKLY_EXPORT_HEADER_LABELS)
	try:
		file_bytes = add_weekly_branding_cover_sheet(file_bytes, summary=summary)
	except Exception:
		pass
	response = HttpResponse(
		file_bytes,
		content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
	)
	response['Content-Disposition'] = f'attachment; filename="{week_report_filename("xlsx", summary.get("processing_date"))}"'
	return response


@require_GET
def download_weekly_csv(request: HttpRequest):
	rows, summary = _weekly_session_data(request)
	if not rows:
		messages.error(request, 'No processed data available. Upload files first.')
		return redirect('weekly:weekly_report')
	file_bytes = build_weekly_csv_bytes(rows, summary=summary)
	response = HttpResponse(file_bytes, content_type='text/csv; charset=utf-8')
	response['Content-Disposition'] = f'attachment; filename="{week_report_filename("csv", summary.get("processing_date"))}"'
	return response


@require_GET
def download_weekly_pdf(request: HttpRequest):
	rows, summary = _weekly_session_data(request)
	if not rows:
		messages.error(request, 'No processed data available. Upload files first.')
		return redirect('weekly:weekly_report')
	file_bytes = build_weekly_pdf_bytes(rows, summary=summary)
	response = HttpResponse(file_bytes, content_type='application/pdf')
	response['Content-Disposition'] = f'attachment; filename="{week_report_filename("pdf", summary.get("processing_date"))}"'
	return response


@require_GET
def download_monthly_excel(request: HttpRequest):
	blob = request.session.get('monthly_last', {})
	raw = blob.get('summaries_json')
	if not raw:
		messages.error(request, 'No monthly data. Upload weekly files first.')
		return redirect('weekly:monthly_report')
	try:
		summaries = monthly_summaries_from_json(raw)
		file_bytes = build_monthly_excel_bytes(summaries)
	except Exception as exc:
		messages.error(request, f'Could not build Excel: {exc}')
		return redirect('weekly:monthly_report')
	response = HttpResponse(
		file_bytes,
		content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
	)
	date_label = month_report_date_label(summaries)
	response['Content-Disposition'] = f'attachment; filename="{month_report_filename("xlsx", date_label)}"'
	return response
