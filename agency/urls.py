from django.urls import path

from agency import views
from agency.views import AgencyUpload, AgencyListView, AgencyDetailsView, AgencySalesDetailsView, GetAgencyList, \
    AgencyTypeView, AgencyTypeDetailsView, AgencyTypeDelete, AgencyTypeCreate, AgencyUpdateView, AgencyTypeUpdateView, \
    AgencyTypeRemoveAgency, StatusHistoryView, AgencyCollectionView, AgencyCollectionDelete, \
    AgencyCollectionDetailsView, AgencyCollectionRemoveAgency, AgencyCollectionCreate, AgencyCollectionUpdateView, \
    AgencyCollectionReportView, AgencyCollectionReportDownloadView

urlpatterns = [
    path('', AgencyListView.as_view(), name='agencies'),
    path('<int:pk>/', AgencyDetailsView.as_view(), name='agency_details'),
    path('<int:pk>/sales/', AgencySalesDetailsView.as_view(), name='agency_sales_details'),
    path('<int:pk>/sales-exel/', views.AgencySalesExelView.as_view(), name='agency_sales_exel'),
    path('<int:pk>/update', AgencyUpdateView.as_view(),
         name='agency_details_update'),

    path('uploads/', AgencyUpload.as_view(), name='agency_uploads'),
    path('download/', GetAgencyList.as_view(), name='agency_download'),

    path('types/', AgencyTypeView.as_view(), name='agency_types'),
    path('types/<int:pk>/', AgencyTypeDetailsView.as_view(),
         name='agency_type_details'),
    path('types/<int:pk>/delete', AgencyTypeDelete.as_view(),
         name='agency_type_delete'),
    path('types/<int:pk>/update', AgencyTypeUpdateView.as_view(),
         name='agency_type_update'),
    path('types/create', AgencyTypeCreate.as_view(),
         name='agency_type_create'),
    path('types/removeagency', AgencyTypeRemoveAgency.as_view(),
         name='agency_type_remove'),

    path('collections/', AgencyCollectionView.as_view(), name='agency_collections'),
    path('collections/<int:pk>/', AgencyCollectionDetailsView.as_view(),
         name='agency_collection_details'),
    path('collections/<int:pk>/delete', AgencyCollectionDelete.as_view(),
         name='agency_collection_delete'),
    path('collections/<int:pk>/update', AgencyCollectionUpdateView.as_view(),
         name='agency_collection_update'),
    path('collections/create', AgencyCollectionCreate.as_view(),
         name='agency_collection_create'),
    path('collections/removeagency', AgencyCollectionRemoveAgency.as_view(),
         name='agency_collection_remove'),
    path('collections/<int:pk>/report', AgencyCollectionReportView.as_view(),
         name='agency_collection_report'),
    path('collections/<int:pk>/report-download', AgencyCollectionReportDownloadView.as_view(),
         name='agency_collection_report_download'),

    path('<int:pk>/history', StatusHistoryView.as_view(), name='status_history'),
    path('state-owners/', views.StateOwnersListView.as_view(),
         name='state_owners'),
    path('state-owners/<int:pk>/update/',
         views.StateOwnerUpdateView.as_view(), name='update_state_owner'),

    # path('users/<int:pk>/update/',
    #      views.UserUpdateView.as_view(), name='update_user'),
    # path('users/<int:pk>/delete/',
    #      views.UserDeleteView.as_view(), name='delete_user'),

     path('new-agent-list/', views.NewAgentReportView.as_view(),
         name='new_agent_report'),
     path('get-new-agent-report/', views.NewAgentReportDownloadView.as_view(),
         name='get_new_agent_report'),
]
