from __future__ import annotations

from collections import defaultdict
from typing import Any

from .payroll_service import _HOUR_BAND_COLS, _sum_hour_bands, split_emp_agency_rows


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

	gazebo_rows, agency_rows = split_emp_agency_rows(rows)
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
