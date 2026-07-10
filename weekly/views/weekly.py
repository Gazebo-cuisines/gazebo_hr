from __future__ import annotations

from django.contrib import messages
from django.http import HttpRequest
from django.shortcuts import redirect, render
from django.views.decorators.http import require_GET, require_http_methods

from ..analytics import weekly_analytics_from_rows
from ..payroll_service import (
	audit_contract_integrity,
	calculate_weekly_payroll,
	parse_employee_hours,
	parse_processing_date,
)
from .common import _contract_audit_from_session


@require_GET
def weekly_help(request: HttpRequest):
	return render(
		request,
		'weekly/help_weekly.html',
		{
			'title': 'Weekly report — How to use',
			'page_heading': 'Weekly report — How to use',
		},
	)


@require_http_methods(['GET', 'POST'])
def weekly_report(request: HttpRequest):
	result_data = request.session.get('weekly_last_result', {})
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
			return redirect('weekly:weekly_report')
		try:
			employee_file.seek(0)
			processing_date = parse_processing_date(employee_file)
			employee_file.seek(0)
			employee_rows = parse_employee_hours(employee_file)
			payroll_result = calculate_weekly_payroll(employee_rows, contracted_file)
			contracted_file.seek(0)
			audit = audit_contract_integrity(employee_rows, contracted_file, payroll_result.rows)
		except Exception as exc:
			messages.error(request, f'Could not process files: {exc}')
			return redirect('weekly:weekly_report')

		request.session['weekly_last_result'] = {
			'rows': payroll_result.rows,
			'summary': {
				'total_rows': len(payroll_result.rows),
				'total_paid_hours': payroll_result.total_paid_hours,
				'agency_rows': len(payroll_result.agency_rows),
				'gazebo_rows': len(payroll_result.gazebo_rows),
				'processing_date': processing_date,
			},
			'contract_audit': {
				'missing': audit.missing,
				'conflicts': audit.conflicts,
				'review': audit.review,
			},
		}
		request.session.modified = True
		messages.success(request, f'Processed {len(payroll_result.rows)} employee rows.')
		return redirect('weekly:weekly_report')

	return render(
		request,
		'weekly/weekly_report.html',
		{
			'title': 'Weekly report — Gazebo',
			'page_heading': 'Weekly report',
			'preview_rows': preview_rows,
			'summary': summary,
			'contract_audit': contract_audit,
			'weekly_analytics': weekly_analytics,
		},
	)


@require_http_methods(['POST'])
def weekly_clear_results(request: HttpRequest):
	request.session.pop('weekly_last_result', None)
	request.session.modified = True
	messages.success(request, 'Previous results cleared.')
	return redirect('weekly:weekly_report')
