from django.urls import path

from .views import auth, case_studies, common, daily, downloads, monthly, weekly

app_name = 'weekly'

urlpatterns = [
	path('', auth.home, name='home'),
	path('login/', auth.login_view, name='login'),
	path('logout/', auth.logout_view, name='logout'),
	path('dashboard/', auth.dashboard, name='dashboard'),
	path('dashboard/daily-report/', daily.daily_report, name='daily_report'),
	path('dashboard/daily-report/clear/', daily.daily_clear_results, name='daily_clear_results'),
	path('dashboard/daily-report/help/', daily.daily_help, name='daily_help'),
	path('dashboard/daily-report/download/', downloads.download_daily_excel, name='daily_download'),
	path('dashboard/daily-report/download.csv', downloads.download_daily_csv, name='daily_download_csv'),
	path('dashboard/daily-report/download.pdf', downloads.download_daily_pdf, name='daily_download_pdf'),

	# Weekly report urls
	path('dashboard/weekly-report/', weekly.weekly_report, name='weekly_report'),
	path('dashboard/weekly-report/clear/', weekly.weekly_clear_results, name='weekly_clear_results'),
	path('dashboard/weekly-report/help/', weekly.weekly_help, name='weekly_help'),
	path('dashboard/weekly-report/download/', downloads.download_weekly_excel, name='weekly_download'),
	path('dashboard/weekly-report/download.csv', downloads.download_weekly_csv, name='weekly_download_csv'),
	path('dashboard/weekly-report/download.pdf', downloads.download_weekly_pdf, name='weekly_download_pdf'),

	# Case studies urls
	path('dashboard/case-studies/', case_studies.case_studies, name='case_studies'),
	path('dashboard/case-studies/<slug:case_id>/', case_studies.case_study_detail, name='case_study_detail'),

	# Monthly report urls
	path('dashboard/monthly-report/', monthly.monthly_report, name='monthly_report'),
	path('dashboard/monthly-report/employee-hour-contracts/', monthly.employee_hour_contracts, name='employee_hour_contracts'),
	path('dashboard/monthly-report/download/', downloads.download_monthly_excel, name='monthly_download'),

	# Health api url
	path('api/health', common.health_api, name='health'),
]
