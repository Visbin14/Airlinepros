from django.urls import path

from report import views

urlpatterns = [


    path('BSP-Sales Summary/',views.NewSalesSummaryReport.as_view(), name='new_sales_summary'),
    path('get BSP-Sales Summary/', views.NewGetSalesSummaryReport.as_view(), name = 'newsummary_report_download'),

    
    path('upload/', views.ReportUpload.as_view(), name='upload_report'),
    path('scheduler-report-upload/', views.SchedulerReportUpload.as_view(), name='scheduler_report_upload'),

    path('sales/', views.SalesReport.as_view(), name='sales_report'),
    path('sales-by/', views.SalesByReport.as_view(), name='sales_by_report'),
    path('allsales/', views.AllSalesReport.as_view(), name='all_sales_report'),
    path('sales-comparison/', views.SalesComparisonReport.as_view(), name='sales_comparison_report'),
    path('year-year-sales/', views.YearToYearSalesReport.as_view(), name='year_to_year_sales_report'),
    path('commission', views.CommissionReport.as_view(), name='commission_report'),
    path('top-agent', views.TopAgentReport.as_view(), name='top_agent_report'),

    path('summary/', views.SalesSummaryReport.as_view(), name='sales_summary'),
    path('monthly-yoy/', views.MonthlyYOYReport.as_view(), name='monthly_yoy'),
    path('airline-agency/', views.AirlineAgencyReport.as_view(), name='airline_agency'),

    path('adm_ignorelist/', views.adm_report_ignorelist, name='adm_ignore'),
    path('adm/', views.ADMReport.as_view(), name='adm_report'),

    path('disbursement-summary/', views.DisbursementSummary.as_view(), name='disbursement_summary_report'),

    path('getadm/', views.GetADMReport.as_view(), name='adm_report_download'),
    path('getairlineagency/', views.GetAirlineAgencyReport.as_view(), name='airline_agency_download'),
    path('getmonthlyyoy/', views.GetMonthlyYOYReport.as_view(), name='monthly_yoy_download'),
    path('getsummary/', views.GetSalesSummaryReport.as_view(), name='summary_report_download'),
    path('getcomparison/', views.GetSalesComparisonReport.as_view(), name='sales_comparison_report_download'),

    path('getdetails/', views.GetSalesReport.as_view(), name='sales_deatails_report_download'),
    path('get-sales-by/', views.GetSalesByReport.as_view(), name='sales_by_report_download'),
    path('get-all-sales/', views.GetAllSalesReport.as_view(), name='all_sales_report_download'),
    path('get-yeartoyear-sales/', views.GetYearToYearSalesReport.as_view(), name='year_to_year_sales_report_download'),
    path('get-commission/', views.GetCommissionReport.as_view(), name='commission_report_download'),
    path('get-top-agent/', views.GetTopAgentReport.as_view(), name='top_agent_report_download'),
    path('taxes-partial/<int:pk>/', views.TaxesPartial.as_view(), name='taxes_partial'),
    path('re-process/', views.ReProcessReports.as_view(), name='re_process'),
    path('check-process/', views.CheckTasks.as_view(), name='check_process'),
    path('get-disbursement-summary/', views.GetDisbursementSummary.as_view(), name='disbursement_summary_report_download'),
    path('monthly-airline-summary-report/', views.WeeklySalesReport.as_view(), name='weekly_sales_report'),
    path('get-weekly-sales/', views.GetWeeklySalesReport.as_view(), name='weekly_sales_report_download'),


]
