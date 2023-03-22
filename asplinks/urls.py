"""asplinks URL Configuration

The `urlpatterns` list routes URLs to views. For more information please see:
    https://docs.djangoproject.com/en/2.1/topics/http/urls/
Examples:
Function views
    1. Add an import:  from my_app import views
    2. Add a URL to urlpatterns:  path('', views.home, name='home')
Class-based views
    1. Add an import:  from other_app.views import Home
    2. Add a URL to urlpatterns:  path('', Home.as_view(), name='home')
Including another URLconf
    1. Import the include() function: from django.urls import include, path
    2. Add a URL to urlpatterns:  path('blog/', include('blog.urls'))
"""
from django.conf import settings
from django.conf.urls.static import static
from django.contrib import admin
from django.urls import path, re_path, include
import debug_toolbar
from django.views.generic.base import RedirectView

from report import views

urlpatterns = [
    path('admin/', admin.site.urls),
    path('', include('account.urls')),
    path('', include('main.urls')),
    path('dashboard/', include('Dashboard.urls')),
    path('favicon\.ico', RedirectView.as_view(url='/static/main/img/favicon.ico')),
    path('agencies/', include('agency.urls')),
    path('reports/', include('report.urls')),
    # path('settings/', include('report.urls')),
    path('settings/upload-calendar/', views.CalendarUpload.as_view(), name='upload_calendar'),
    path('settings/calendar/', views.CalendarList.as_view(), name='calendar'),
    path('select2/', include('django_select2.urls')),
] + static(settings.STATIC_URL, document_root=settings.STATIC_ROOT) + static(settings.MEDIA_URL,
                                                                             document_root=settings.MEDIA_ROOT)


if settings.DEBUG:
    urlpatterns = [
        path('__debug__/', include(debug_toolbar.urls)),
    ] + urlpatterns
