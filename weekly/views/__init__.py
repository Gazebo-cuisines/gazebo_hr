"""HTTP views package. urls.py may import callables from this package root."""

from .auth import dashboard, home, login_view, logout_view
from .common import health_api
from .daily import daily_clear_results, daily_help, daily_report
from .weekly import weekly_clear_results, weekly_help, weekly_report
from ._rest import (
	case_studies,
	case_study_detail,
	download_daily_csv,
	download_daily_excel,
	download_daily_pdf,
	download_monthly_excel,
	download_weekly_csv,
	download_weekly_excel,
	download_weekly_pdf,
	employee_hour_contracts,
	monthly_report,
)

__all__ = [
	'case_studies',
	'case_study_detail',
	'daily_clear_results',
	'daily_help',
	'daily_report',
	'dashboard',
	'download_daily_csv',
	'download_daily_excel',
	'download_daily_pdf',
	'download_monthly_excel',
	'download_weekly_csv',
	'download_weekly_excel',
	'download_weekly_pdf',
	'employee_hour_contracts',
	'health_api',
	'home',
	'login_view',
	'logout_view',
	'monthly_report',
	'weekly_clear_results',
	'weekly_help',
	'weekly_report',
]
