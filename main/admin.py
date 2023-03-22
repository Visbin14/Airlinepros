from django.contrib import admin

from .models import *

# Register your models here.
admin.site.register(Airline)
admin.site.register(Airline_Contact)
admin.site.register(Country)
admin.site.register(State)
admin.site.register(City)
admin.site.register(CommissionHistory)
admin.site.register(RemoteServers)
admin.site.register(LatestFiles)
admin.site.register(FTPhistory)
