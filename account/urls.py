from django.urls import path, include
from . import views


urlpatterns = [
    path('', views.CustomLogin.as_view(), name='login'),
    path('login/', views.CustomLogin.as_view(), name='signin'),
    path('password_reset/', views.CustomPasswordReset.as_view(),
         name='password_reset'),
    path('', include('django.contrib.auth.urls')),
    path('users/', views.UserListView.as_view(), name='users'),
    path('users/<int:pk>/', views.UserDetailsView.as_view(), name='user_details'),
    path('users/add/', views.UserCreateView.as_view(), name='add_user'),
    path('users/<int:pk>/update/',
         views.UserUpdateView.as_view(), name='update_user'),
    # path('users/<int:pk>/delete/',
    #      views.UserDeleteView.as_view(), name='delete_user'),
    path('users/<int:pk>/password-reset/',
         views.UserPasswordResetView.as_view(), name='user_password_reset'),
    path('activate/<str:uidb64>/<str:token>/',
         views.activate, name='activate'),
    path('roles/', views.RoleListView.as_view(), name='roles'),
    path('roles/add/', views.RoleCreateView.as_view(), name='add_role'),
    path('roles/<int:pk>/',
         views.RoleDetailView.as_view(), name='role_details'),
    path('roles/<int:pk>/update/',
         views.RoleUpdateView.as_view(), name='update_role'),
    path('roles/<int:pk>/delete/',
         views.RoleDeleteView.as_view(), name='delete_role'),
    path('roles/name-check/',views.get_role_name_status, name='role_name_status')
]
