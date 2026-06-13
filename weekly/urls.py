from django.urls import path

from . import views

app_name = 'weekly'

urlpatterns = [
	path('', views.home, name='home'),
	path('login/', views.login_view, name='login'),
	path('logout/', views.logout_view, name='logout'),
	path('dashboard/', views.dashboard, name='dashboard'),
	path('dashboard/daily-report/', views.daily_report, name='daily_report'),
	path('dashboard/daily-report/clear/', views.daily_clear_results, name='daily_clear_results'),
	path('dashboard/daily-report/help/', views.daily_help, name='daily_help'),
	path('dashboard/daily-report/download/', views.download_daily_excel, name='daily_download'),
	path('dashboard/daily-report/download.csv', views.download_daily_csv, name='daily_download_csv'),
	path('dashboard/daily-report/download.pdf', views.download_daily_pdf, name='daily_download_pdf'),
	path('dashboard/weekly-report/', views.weekly_report, name='weekly_report'),
	path('dashboard/weekly-report/clear/', views.weekly_clear_results, name='weekly_clear_results'),
	path('dashboard/weekly-report/help/', views.weekly_help, name='weekly_help'),
	path('dashboard/weekly-report/download/', views.download_weekly_excel, name='weekly_download'),
	path('dashboard/weekly-report/download.csv', views.download_weekly_csv, name='weekly_download_csv'),
	path('dashboard/weekly-report/download.pdf', views.download_weekly_pdf, name='weekly_download_pdf'),
	path('dashboard/case-studies/', views.case_studies, name='case_studies'),
	path('dashboard/case-studies/<slug:case_id>/', views.case_study_detail, name='case_study_detail'),
	path('dashboard/monthly-report/', views.monthly_report, name='monthly_report'),
	path(
		'dashboard/employee-hour-contracts/',
		views.employee_hour_contracts,
		name='employee_hour_contracts',
	),
	path('dashboard/monthly-report/download/', views.download_monthly_excel, name='monthly_download'),
	path('api/health', views.health_api, name='health'),
]
