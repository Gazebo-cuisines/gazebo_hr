from __future__ import annotations

from typing import Any

from django.http import HttpRequest, JsonResponse
from django.views.decorators.http import require_GET


def _contract_audit_from_session(result_data: dict[str, Any]) -> dict[str, Any]:
	return result_data.get('contract_audit') or {'missing': [], 'conflicts': [], 'review': []}


@require_GET
def health_api(request: HttpRequest):
	return JsonResponse({'ok': True, 'app': 'weekly'})
