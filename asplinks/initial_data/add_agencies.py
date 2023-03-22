
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

# returns JSON object as
# a dictionary
agencies_json_data = json.load(f)
auth_user_json_data = json.load(f2)
agency_list = agencies_json_data.get("RECORDS")
auth_user_list = auth_user_json_data.get("auth_user")
cr_num = 0
up_num = 0
print("STARTED...")
import datetime
today = datetime.datetime.now().date()
print(Agency.objects.get(agency_no=14709995).id)
for ss in City.objects.filter(name="SAN FRANCISCO"):
    print(ss.id)
raise Exception
for data in agency_list:
    agency_no = data.get("agency_no")
    trade_name = data.get("trade_name")
    address1 = data.get("address1")
    address2 = data.get("address2")
    city = data.get("city")
    state = data.get("state")
    zipcode = data.get("zip")
    tel = data.get("tel")
    home_agency_code = data.get("home_agency_code")
    host_agency_code = data.get("host_agency_code")
    email = data.get("email")
    sales_owner_id = data.get("sales_owner_id")
    agency_type = data.get("agency_type")
    for auth_user in auth_user_list:
        if auth_user['id'] == sales_owner_id:
            user_email = auth_user.get("email")
            # store to Agencies model
            try:
                if not Agency.objects.filter(agency_no=agency_no):
                    agency_obj = Agency.objects.create(agency_no=agency_no,
                                                  trade_name=trade_name,
                                                  address1=address1,
                                                  address2=address2,
                                                  from_old_asplinks=True
                                                  )
                    if User.objects.filter(email=user_email):
                        if len(User.objects.filter(email=user_email)) > 1:
                            user = User.objects.filter(email=user_email)[0]
                        else:
                            user = User.objects.get(email=user_email)

                    if City.objects.filter(name=city):
                        if len(City.objects.filter(name=city)) > 1:
                            city = City.objects.filter(name=city)[0]
                        else:
                            city = City.objects.get(name=city)

                        agency_obj.city = city
                        agency_obj.state = city.state
                        agency_obj.country = city.country
                    elif State.objects.filter(name=state):
                        if len(State.objects.filter(name=state))>1:
                            state = State.objects.filter(name=state)[0]
                        else:   
                            state = State.objects.get(name=state)
                        agency_obj.state = state
                        agency_obj.country = state.country

                    if AgencyType.objects.filter(name=agency_type):
                        if len(AgencyType.objects.filter(name=agency_type)) > 1:
                            agency_obj.agency_type = AgencyType.objects.filter(name=agency_type)[0]
                        else:
                            agency_obj.agency_type = AgencyType.objects.get(name=agency_type)

                    agency_obj.zip_code = zipcode
                    agency_obj.email = email
                    agency_obj.tel = tel
                    agency_obj.home_agency = home_agency_code
                    agency_obj.save()
                    print("AGENCY CREATED : ",agency_obj.agency_no)
                    up_num += 1
            except Exception as e:
                print("Error in AGENCY : ",agency_no,"  : ",e)
            
print(up_num," ROWS CREATED")
print("COMPLETE...")