from typing import Any

from django.contrib.auth import authenticate, login, logout
from django.contrib import messages
from django.http import HttpRequest, HttpResponse, JsonResponse
from django.shortcuts import redirect, render
from django.utils.http import url_has_allowed_host_and_scheme
from django.views.decorators.http import require_GET, require_http_methods

from .monthly_service import (
	build_monthly_excel_bytes,
	monthly_summaries_from_json,
	monthly_summaries_to_json,
	parse_monthly_inputs,
)
from .payroll_service import (
	AGENCY_CATEGORIES,
	PayrollResult,
	_HOUR_BAND_COLS,
	_sum_hour_bands,
	build_excel_bytes,
	calculate_payroll,
	parse_employee_hours,
	total_paid_hours_from_rows,
)
from .case_studies_data import CASE_STUDIES, get_case_study
from .export_service import (
	add_branding_cover_sheet,
	build_csv_bytes,
	build_pdf_bytes,
	export_filename,
)


def _hours_to_float(value: Any) -> float:
	if value is None:
		return 0.0
	if isinstance(value, (int, float)):
		return float(value)
	s = str(value).strip().replace(',', '')
	if not s:
		return 0.0
	try:
		return float(s)
	except ValueError:
		return 0.0


def _rollup_categories(rows: list[dict[str, Any]], top_n: int = 10) -> tuple[list[str], list[float], list[int]]:
	from collections import defaultdict

	hours: dict[str, float] = defaultdict(float)
	counts: dict[str, int] = defaultdict(int)
	for row in rows:
		cat = str(row.get('Category') or '').strip() or '(none)'
		hours[cat] += _hours_to_float(row.get('TotalPaidHours'))
		counts[cat] += 1
	ordered = sorted(hours.keys(), key=lambda c: hours[c], reverse=True)
	if len(ordered) <= top_n:
		labels = ordered
		h_vals = [round(hours[c], 2) for c in labels]
		c_vals = [counts[c] for c in labels]
		return labels, h_vals, c_vals
	top = ordered[:top_n]
	rest = ordered[top_n:]
	h_other = sum(hours[c] for c in rest)
	c_other = sum(counts[c] for c in rest)
	labels = top + ['Other']
	h_vals = [round(hours[c], 2) for c in top] + [round(h_other, 2)]
	c_vals = [counts[c] for c in top] + [c_other]
	return labels, h_vals, c_vals


def weekly_analytics_from_rows(rows: list[dict[str, Any]]) -> dict[str, Any] | None:
	if not rows:
		return None
	buckets = [0, 0, 0, 0]
	over_60: list[dict[str, Any]] = []
	for row in rows:
		h = _hours_to_float(row.get('TotalPaidHours'))
		if h < 40:
			buckets[0] += 1
		elif h < 48:
			buckets[1] += 1
		elif h < 60:
			buckets[2] += 1
		else:
			buckets[3] += 1
		if h >= 60.0:
			over_60.append(
				{
					'Name': row.get('Name') or '',
					'SageNo': row.get('SageNo'),
					'Category': row.get('Category') or '',
					'TotalPaidHours': h,
				}
			)
	over_60.sort(key=lambda x: x['TotalPaidHours'], reverse=True)

	gazebo_rows = [r for r in rows if str(r.get('Category', '')).strip().upper() not in AGENCY_CATEGORIES]
	agency_rows = [r for r in rows if str(r.get('Category', '')).strip().upper() in AGENCY_CATEGORIES]
	emp_totals = _sum_hour_bands(gazebo_rows)
	ag_totals = _sum_hour_bands(agency_rows)
	cat_labels, cat_hours, cat_counts = _rollup_categories(rows)
	_palette = [
		'#005ea5',
		'#85994b',
		'#f47738',
		'#528187',
		'#7d4b8c',
		'#b10e1e',
		'#ffbf47',
		'#5694ca',
		'#67874e',
		'#f499be',
		'#505a5f',
	]
	_colors = [_palette[i % len(_palette)] for i in range(len(cat_labels))]

	return {
		'total_people': len(rows),
		'over_60_count': len(over_60),
		'over_60': over_60[:200],
		'chart': {
			'labels': ['Under 40 h', '40–48 h', '48–60 h', '60+ h'],
			'counts': buckets,
		},
		'extra_charts': {
			'category': {
				'labels': cat_labels,
				'hours': cat_hours,
				'counts': cat_counts,
				'colors': _colors,
			},
			'empAgency': {
				'bandLabels': ['Basic', 'Mon–Fri OT', 'Sat–Sun OT', 'Annual', 'Total paid'],
				'emp': [round(emp_totals[k], 2) for k in _HOUR_BAND_COLS],
				'agency': [round(ag_totals[k], 2) for k in _HOUR_BAND_COLS],
			},
			'totalPaidSplit': {
				'emp': round(float(emp_totals['TotalPaidHours']), 2),
				'agency': round(float(ag_totals['TotalPaidHours']), 2),
			},
		},
	}


@require_GET
def home(request: HttpRequest):
	if request.user.is_authenticated:
		return redirect('weekly:weekly_report')
	return render(
		request,
		'weekly/home.html',
		{
			'title': 'Gazebo',
			'page_heading': 'Payroll and reporting',
		},
	)


@require_http_methods(['GET', 'POST'])
def login_view(request: HttpRequest):
	if request.user.is_authenticated:
		return redirect('weekly:dashboard')
	error = ''
	if request.method == 'POST':
		username = request.POST.get('username', '').strip()
		password = request.POST.get('password', '')
		user = authenticate(request, username=username, password=password)
		if user is not None:
			login(request, user)
			next_url = (request.POST.get('next') or request.GET.get('next') or '').strip()
			if next_url and url_has_allowed_host_and_scheme(
				next_url,
				allowed_hosts={request.get_host()},
				require_https=request.is_secure(),
			):
				return redirect(next_url)
			return redirect('weekly:dashboard')
		error = 'Enter a valid username and password.'
	return render(
		request,
		'weekly/login.html',
		{
			'title': 'Sign in — Gazebo',
			'page_heading': 'Sign in',
			'error': error,
		},
	)


@require_http_methods(['GET', 'POST'])
def logout_view(request: HttpRequest):
	logout(request)
	return redirect('weekly:home')


@require_GET
def dashboard(request: HttpRequest):
	return render(
		request,
		'weekly/dashboard.html',
		{
			'title': 'Dashboard — Gazebo',
			'page_heading': 'Dashboard',
		},
	)


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
		from django.http import Http404
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


@require_http_methods(['GET', 'POST'])
def weekly_report(request: HttpRequest):
	result_data = request.session.get('weekly_last_result', {})
	all_rows = result_data.get('rows', [])
	preview_rows = all_rows[:200]
	summary = result_data.get('summary', {})
	weekly_analytics = weekly_analytics_from_rows(all_rows)

	if request.method == 'POST':
		employee_file = request.FILES.get('employee_file')
		contracted_file = request.FILES.get('contracted_file')
		if not employee_file or not contracted_file:
			messages.error(request, 'Upload both files: employee hours and contracted hours.')
			return redirect('weekly:weekly_report')
		try:
			employee_rows = parse_employee_hours(employee_file)
			payroll_result = calculate_payroll(employee_rows, contracted_file)
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
			'weekly_analytics': weekly_analytics,
		},
	)


@require_http_methods(['POST'])
def weekly_clear_results(request: HttpRequest):
	request.session.pop('weekly_last_result', None)
	request.session.modified = True
	messages.success(request, 'Previous results cleared.')
	return redirect('weekly:weekly_report')


def _parse_non_hourly_names(raw: str) -> set[str]:
	names: set[str] = set()
	for line in (raw or '').splitlines():
		n = line.strip()
		if n:
			names.add(n.upper())
	return names


def _monthly_context(session_blob: dict) -> dict:
	raw = session_blob.get('summaries_json') or []
	summaries = monthly_summaries_from_json(raw) if raw else []
	week_cards: list[dict[str, Any]] = []
	merged_totals: dict[str, dict[str, float]] = {}
	merged_grouped: dict[str, dict[str, float]] = {}
	total_paid_hours = 0.0
	total_employee_rows = 0
	total_non_agency = 0.0
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

	return {
		'week_cards': week_cards,
		'week_count': len(summaries),
		'non_hourly_names': session_blob.get('non_hourly') or [],
		'category_totals': category_totals,
		'grouped_totals': grouped_totals,
		'summary_stats': {
			'total_paid_hours': round(total_paid_hours, 2),
			'non_agency_total': round(total_non_agency, 2),
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
			summaries = parse_monthly_inputs(files)
			non_hourly = _parse_non_hourly_names(request.POST.get('non_hourly_names', ''))
			for s in summaries:
				for e in s.employees:
					if e.Name.strip().upper() in non_hourly and e.Category.strip().startswith('A-'):
						messages.error(request, f'Agency employee must be hourly: {e.Name}')
						return redirect('weekly:monthly_report')
			request.session['monthly_last'] = {
				'summaries_json': monthly_summaries_to_json(summaries),
				'non_hourly': sorted(non_hourly),
				'week_count': len(summaries),
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
		non_hourly = set(blob.get('non_hourly') or [])
		file_bytes = build_monthly_excel_bytes(summaries, non_hourly_names=non_hourly)
	except Exception as exc:
		messages.error(request, f'Could not build Excel: {exc}')
		return redirect('weekly:monthly_report')
	response = HttpResponse(
		file_bytes,
		content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
	)
	response['Content-Disposition'] = 'attachment; filename="monthly_report_output.xlsx"'
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


@require_GET
def health_api(request: HttpRequest):
	return JsonResponse({'ok': True, 'app': 'weekly'})


def _weekly_session_data(request: HttpRequest) -> tuple[list[dict], dict]:
	result_data = request.session.get('weekly_last_result', {})
	return result_data.get('rows', []), result_data.get('summary', {})


@require_GET
def download_weekly_excel(request: HttpRequest):
	rows, summary = _weekly_session_data(request)
	if not rows:
		messages.error(request, 'No processed data available. Upload files first.')
		return redirect('weekly:weekly_report')

	agency_rows = [r for r in rows if str(r.get('Category', '')).strip().upper() in AGENCY_CATEGORIES]
	gazebo_rows = [r for r in rows if str(r.get('Category', '')).strip().upper() not in AGENCY_CATEGORIES]
	payroll_result = PayrollResult(
		rows=rows,
		agency_rows=agency_rows,
		gazebo_rows=gazebo_rows,
		total_paid_hours=total_paid_hours_from_rows(rows),
	)
	file_bytes = build_excel_bytes(payroll_result)
	try:
		file_bytes = add_branding_cover_sheet(file_bytes, summary=summary)
	except Exception:
		pass
	response = HttpResponse(
		file_bytes,
		content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
	)
	response['Content-Disposition'] = f'attachment; filename="{export_filename("gazebo_weekly_report", "xlsx")}"'
	return response


@require_GET
def download_weekly_csv(request: HttpRequest):
	rows, summary = _weekly_session_data(request)
	if not rows:
		messages.error(request, 'No processed data available. Upload files first.')
		return redirect('weekly:weekly_report')
	file_bytes = build_csv_bytes(rows, summary=summary)
	response = HttpResponse(file_bytes, content_type='text/csv; charset=utf-8')
	response['Content-Disposition'] = f'attachment; filename="{export_filename("gazebo_weekly_report", "csv")}"'
	return response


@require_GET
def download_weekly_pdf(request: HttpRequest):
	rows, summary = _weekly_session_data(request)
	if not rows:
		messages.error(request, 'No processed data available. Upload files first.')
		return redirect('weekly:weekly_report')
	file_bytes = build_pdf_bytes(rows, summary=summary)
	response = HttpResponse(file_bytes, content_type='application/pdf')
	response['Content-Disposition'] = f'attachment; filename="{export_filename("gazebo_weekly_report", "pdf")}"'
	return response


