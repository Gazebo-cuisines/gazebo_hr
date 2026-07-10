"""Views not yet split into feature modules (Chunk 11)."""

from __future__ import annotations

from typing import Any

from django.contrib import messages
from django.http import Http404, HttpRequest, HttpResponse
from django.shortcuts import redirect, render
from django.views.decorators.http import require_GET, require_http_methods

from ..case_studies_data import CASE_STUDIES, get_case_study
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
from ..monthly_service import (
	build_monthly_excel_bytes,
	monthly_summaries_from_json,
	monthly_summaries_to_json,
	parse_monthly_inputs,
)
from ..payroll_service import (
	PayrollResult,
	build_excel_bytes,
	build_staff_summary,
	split_emp_agency_rows,
	total_paid_hours_from_rows,
)


@require_GET
def case_studies(request: HttpRequest):
	return render(
		request,
		'weekly/case_studies.html',
		{
			'title': 'Case studies — Gazebo',
			'page_heading': 'Case studies',
			'case_studies': CASE_STUDIES,
		},
	)


@require_GET
def case_study_detail(request: HttpRequest, case_id: str):
	case = get_case_study(case_id)
	if case is None:
		raise Http404('Case study not found')
	return render(
		request,
		'weekly/case_study_detail.html',
		{
			'title': f'{case["title"]} — Case studies',
			'page_heading': case['title'],
			'case': case,
		},
	)


def _monthly_context(session_blob: dict) -> dict:
	raw = session_blob.get('summaries_json') or []
	summaries = monthly_summaries_from_json(raw) if raw else []
	week_cards: list[dict[str, Any]] = []
	merged_totals: dict[str, dict[str, float]] = {}
	merged_grouped: dict[str, dict[str, float]] = {}
	total_paid_hours = 0.0
	total_employee_rows = 0
	total_non_agency = 0.0
	total_additional_holiday_pay = 0.0
	for idx, s in enumerate(summaries, start=1):
		week_total = sum(float(t.TotalPaidHours) for t in s.employee_totals)
		week_cards.append({
			'index': idx,
			'start_date': s.start_date,
			'end_date': s.end_date,
			'employee_count': len(s.employees),
			'category_count': len(s.employee_totals),
			'adjustments_count': len(s.adjustments),
			'non_agency_total': round(float(s.non_agency_total or 0.0), 2),
			'total_paid_hours': round(week_total, 2),
		})
		total_paid_hours += week_total
		total_non_agency += float(s.non_agency_total or 0.0)
		total_additional_holiday_pay += float(s.total_additional_holiday_pay or 0.0)
		total_employee_rows += len(s.employees)
		for t in s.employee_totals:
			cur = merged_totals.setdefault(
				t.Category,
				{'Category': t.Category, 'BasicHours': 0.0, 'MonFriOvertime': 0.0, 'SatSunOvertime': 0.0, 'AnnualHoliday': 0.0, 'TotalPaidHours': 0.0},
			)
			cur['BasicHours'] += float(t.BasicHours)
			cur['MonFriOvertime'] += float(t.MonFriOvertime)
			cur['SatSunOvertime'] += float(t.SatSunOvertime)
			cur['AnnualHoliday'] += float(t.AnnualHoliday)
			cur['TotalPaidHours'] += float(t.TotalPaidHours)
		for k, t in (s.grouped_totals or {}).items():
			cur = merged_grouped.setdefault(
				k,
				{'Category': k, 'BasicHours': 0.0, 'MonFriOvertime': 0.0, 'SatSunOvertime': 0.0, 'AnnualHoliday': 0.0, 'TotalPaidHours': 0.0},
			)
			cur['BasicHours'] += float(t.BasicHours)
			cur['MonFriOvertime'] += float(t.MonFriOvertime)
			cur['SatSunOvertime'] += float(t.SatSunOvertime)
			cur['AnnualHoliday'] += float(t.AnnualHoliday)
			cur['TotalPaidHours'] += float(t.TotalPaidHours)

	def _round(d: dict) -> dict:
		return {k: (round(v, 2) if isinstance(v, float) else v) for k, v in d.items()}

	category_totals = sorted((_round(v) for v in merged_totals.values()), key=lambda r: r['TotalPaidHours'], reverse=True)
	grouped_totals = sorted((_round(v) for v in merged_grouped.values()), key=lambda r: r['TotalPaidHours'], reverse=True)

	stored_dates = session_blob.get('week_dates') or []
	week_date_fields = [
		{
			'start': stored_dates[i][0] if i < len(stored_dates) else '',
			'end': stored_dates[i][1] if i < len(stored_dates) else '',
		}
		for i in range(5)
	]

	return {
		'week_cards': week_cards,
		'week_count': len(summaries),
		'week_date_fields': week_date_fields,
		'category_totals': category_totals,
		'grouped_totals': grouped_totals,
		'summary_stats': {
			'total_paid_hours': round(total_paid_hours, 2),
			'non_agency_total': round(total_non_agency, 2),
			'total_additional_holiday_pay': round(total_additional_holiday_pay, 2),
			'employee_rows': total_employee_rows,
		},
	}


@require_http_methods(['GET', 'POST'])
def monthly_report(request: HttpRequest):
	if request.method == 'POST':
		files_in_order = [
			request.FILES.get('week1'),
			request.FILES.get('week2'),
			request.FILES.get('week3'),
			request.FILES.get('week4'),
			request.FILES.get('week5'),
		]
		use5 = bool(request.POST.get('use_week5'))
		required = 5 if use5 else 4
		files = [f for f in files_in_order[:required] if f]
		if len(files) < required:
			messages.error(request, f'Upload {required} weekly workbook(s) (xls/xlsx), or enable week 5 and upload all five.')
			return redirect('weekly:monthly_report')
		try:
			week_dates = []
			for n in ('1', '2', '3', '4', '5'):
				start = (request.POST.get(f'week{n}_start') or '').strip()
				end = (request.POST.get(f'week{n}_end') or '').strip()
				week_dates.append((start, end))
			summaries = parse_monthly_inputs(files, week_dates=week_dates)
			for wi, s in enumerate(summaries, start=1):
				if not s.employees:
					messages.error(
						request,
						f'Week {wi}: no employees found in All Data — check file format.',
					)
					return redirect('weekly:monthly_report')
			request.session['monthly_last'] = {
				'summaries_json': monthly_summaries_to_json(summaries),
				'week_count': len(summaries),
				'week_dates': week_dates[: len(summaries)],
			}
			request.session.modified = True
			messages.success(request, f'Processed {len(summaries)} weekly file(s). Download Excel when ready.')
		except Exception as exc:
			messages.error(request, f'Could not process monthly files: {exc}')
			return redirect('weekly:monthly_report')
		return redirect('weekly:monthly_report')

	session_blob = request.session.get('monthly_last', {})
	ctx = {
		'title': 'Monthly report — Gazebo',
		'page_heading': 'Monthly report',
	}
	ctx.update(_monthly_context(session_blob))
	return render(request, 'weekly/monthly_report.html', ctx)


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


@require_GET
def employee_hour_contracts(request: HttpRequest):
	return render(
		request,
		'weekly/employee_hour_contracts.html',
		{
			'title': 'Employee hour contracts — Gazebo',
			'page_heading': 'Employee hour contracts',
		},
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


def _weekly_session_data(request: HttpRequest) -> tuple[list[dict], dict]:
	result_data = request.session.get('weekly_last_result', {})
	return result_data.get('rows', []), result_data.get('summary', {})


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
