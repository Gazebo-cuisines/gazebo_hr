from __future__ import annotations

from django.http import Http404, HttpRequest
from django.shortcuts import render
from django.views.decorators.http import require_GET

from ..case_studies_data import CASE_STUDIES, get_case_study


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
