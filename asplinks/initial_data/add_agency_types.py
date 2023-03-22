
import django
import sys
import os



project_dir_path = os.path.abspath(os.getcwd())
sys.path.append(project_dir_path)
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'asplinks.settings')
django.setup()

agencies_json_path = 'asplinks/initial_data/AgencyHandler_agency.json'
auth_user_json_path = 'asplinks/initial_data/auth_user.json'
import json
from account.models import User
from agency.models import *
from main.models import *
# Opening JSON file
f = open(agencies_json_path)
f2 = open(auth_user_json_path)

agencies_json_data = json.load(f)
auth_user_json_data = json.load(f2)
agency_list = agencies_json_data.get("RECORDS")
for data in agency_list:
    agency_no = data.get("agency_no")
    trade_name = data.get("trade_name")
    address1 = data.get("address1")
    address2 = data.get("address2")
    city = data.get("city")
    state = data.get("state")
    zip = data.get("zip")
    tel = data.get("tel")
    home_agency_code = data.get("home_agency_code")
    host_agency_code = data.get("host_agency_code")
    email = data.get("email")
    sales_owner_id = data.get("sales_owner_id")
    agency_type = data.get("agency_type")
    country = None
    if City.objects.filter(name=city):
        if len(City.objects.filter(name=city)) > 1:
            city = City.objects.filter(name=city)[0]
        else:
            city = City.objects.get(name=city)
        country = city.country if city.country.name not in ["None",""] else None
    elif State.objects.filter(name=state):
        if len(State.objects.filter(name=state)) > 1:
            state = State.objects.filter(name=state)[0]
        else:
            state = State.objects.get(name=state)
        country = state.country if state.country.name not in ["None",""] else None
    if agency_type:
        agency_type = agency_type.replace(" ","")
        if agency_type != "None" and not AgencyType.objects.filter(name=agency_type):
            agency_type = AgencyType.objects.create(name=agency_type)
            if country:
                agency_type.country = country
            agency_type.save()
