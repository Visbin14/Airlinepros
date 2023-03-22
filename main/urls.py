from django.urls import path
from . import views


urlpatterns = [
    path('dashboard/', views.HomeView.as_view(), name='home'),
    path('airlines/', views.AirlineListView.as_view(), name='airlines'),
    path('airlines/<int:pk>/', views.AirlineDetailsView.as_view(), name='airline_details'),
    path('airlines/add/', views.AirlineCreateView.as_view(), name='add_airline'),
    path('airlines/<int:pk>/update/',
         views.AirlineUpdateView.as_view(), name='update_airline'),
    path('airlines/<int:pk>/delete/',
         views.AirlineDeleteView.as_view(), name='delete_airline'),
    path('airlines/<int:pk>/commissions/',
         views.AirlineCommissionsView.as_view(), name='airline_commissions'),
    path('airlines/commissions/<int:pk>/delete/', views.AirlineCommissionDelete.as_view(),
         name='airline_commission_delete'),
    path('set-country/', views.SetCountryView.as_view(), name='set_country'),
    path('add-country/', views.AddCountryView.as_view(), name='add_country'),
    path('countries/', views.ListCountryView.as_view(), name='add_country'),
    path('ftp-management/', views.FTPManagementView.as_view(), name='ftp_management'),
    path('add-remote-host/', views.AddRemoteHostView.as_view(), name='add_remote_host'),
    path('country/<int:pk>/update/',
         views.CountryUpdateView.as_view(), name='update_country'),

    ]
