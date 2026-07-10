from __future__ import annotations

from typing import Any

from django.contrib import messages
from django.http import HttpRequest
from django.shortcuts import redirect, render
from django.views.decorators.http import require_GET, require_http_methods

from ..monthly_service import (
	monthly_summaries_from_json,
	monthly_summaries_to_json,
	parse_monthly_inputs,
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
def employee_hour_contracts(request: HttpRequest):
	return render(
		request,
		'weekly/employee_hour_contracts.html',
		{
			'title': 'Employee hour contracts — Gazebo',
			'page_heading': 'Employee hour contracts',
		},
	)
