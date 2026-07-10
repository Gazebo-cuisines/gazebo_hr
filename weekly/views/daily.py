from __future__ import annotations

from django.contrib import messages
from django.http import HttpRequest
from django.shortcuts import redirect, render
from django.views.decorators.http import require_GET, require_http_methods

from ..analytics import weekly_analytics_from_rows
from ..payroll_service import (
	audit_contract_integrity,
	build_staff_summary,
	calculate_payroll,
	parse_employee_hours,
	parse_processing_date,
)
from .common import _contract_audit_from_session


@require_GET
def daily_help(request: HttpRequest):
	return render(
		request,
		'weekly/help_daily.html',
		{
			'title': 'Daily report — How to use',
			'page_heading': 'Daily report — How to use',
		},
	)


@require_http_methods(['GET', 'POST'])
def daily_report(request: HttpRequest):
	result_data = request.session.get('daily_last_result', {})
	all_rows = result_data.get('rows', [])
	preview_rows = all_rows[:200]
	summary = result_data.get('summary', {})
	contract_audit = _contract_audit_from_session(result_data)
	weekly_analytics = weekly_analytics_from_rows(all_rows)

	if request.method == 'POST':
		employee_file = request.FILES.get('employee_file')
		contracted_file = request.FILES.get('contracted_file')
		if not employee_file or not contracted_file:
			messages.error(request, 'Upload both files: employee hours and contracted hours.')
			return redirect('weekly:daily_report')
		try:
			employee_file.seek(0)
			processing_date = parse_processing_date(employee_file)
			employee_file.seek(0)
			employee_rows = parse_employee_hours(employee_file)
			payroll_result = calculate_payroll(employee_rows, contracted_file)
			contracted_file.seek(0)
			audit = audit_contract_integrity(employee_rows, contracted_file, payroll_result.rows)
		except Exception as exc:
			messages.error(request, f'Could not process files: {exc}')
			return redirect('weekly:daily_report')

		staff_summary = build_staff_summary(payroll_result.rows)
		request.session['daily_last_result'] = {
			'rows': payroll_result.rows,
			'summary': {
				**staff_summary,
				'processing_date': processing_date,
				'operator': 'HR',
				'total_rows': len(payroll_result.rows),
				'agency_rows': len(payroll_result.agency_rows),
				'gazebo_rows': len(payroll_result.gazebo_rows),
			},
			'contract_audit': {
				'missing': audit.missing,
				'conflicts': audit.conflicts,
				'review': audit.review,
			},
		}
		request.session.modified = True
		messages.success(request, f'Processed {len(payroll_result.rows)} employee rows.')
		return redirect('weekly:daily_report')

	return render(
		request,
		'weekly/daily_report.html',
		{
			'title': 'Daily report — Gazebo',
			'page_heading': 'Daily report',
			'preview_rows': preview_rows,
			'summary': summary,
			'contract_audit': contract_audit,
			'weekly_analytics': weekly_analytics,
		},
	)


@require_http_methods(['POST'])
def daily_clear_results(request: HttpRequest):
	request.session.pop('daily_last_result', None)
	request.session.modified = True
	messages.success(request, 'Previous results cleared.')
	return redirect('weekly:daily_report')
