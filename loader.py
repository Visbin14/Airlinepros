
import json
from unicodedata import name

import django
import sys
import os

project_dir_path = os.path.abspath(os.getcwd())
sys.path.append(project_dir_path)
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'asplinks.settings')
django.setup()

from agency.models import Agency, AgencyType

def loader():

    print("it is loading......")
    f= open('agencyData')
    data=json.load(f)
    for i in data:
        if i == "NaN" or data[i] == "NaN":
            continue
        else:
            
            obj= Agency.objects.filter(agency_no = i.strip()).first()
            if obj:
                print(obj,"////////")
                try:
                    name="".join(data[i].rstrip().lstrip())
                    if name== "Consolidator":
                        name=name+"s"
                    print(name,"/////")
                    typeobj= AgencyType.objects.get(name= name)
                    print(typeobj)
                    obj.agency_type= typeobj
                    obj.save()
                except:
                    pass
                
            

            #     except Exception as e:
            #         print(e)
            #         pass

            #     print(obj)
            # except:
            #     pass
            

if __name__ == "__main__":
    loader()

