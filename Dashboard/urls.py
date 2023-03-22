from django.contrib import admin
from django.urls import path
from Dashboard import views
from Dashboard.views import Uploaddashboardfile,Dashboard,Pivot


urlpatterns = [
    
    path('airlinedashboard/',Dashboard.as_view(),name='dashboard'),
    path('upload-file/',Uploaddashboardfile.as_view(),name='upload'),
    path('airlinedashboard/pivot',Pivot.as_view(),name='pivot'),
    path('get_Pivot-data/', views.Getpivotdata.as_view(), name = 'Pivot_data_download'),
]