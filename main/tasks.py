from django.template.loader import render_to_string
from django.core.mail import EmailMessage

from report.models import *
from celery import shared_task
from main.models import State, Country


def send_mail(subject, template, context, to, from_email='info@regmat.com'):
    """

    Basic send mail with custom html template
    """
    from_email = 'assda@airlinepros.com'
    message = render_to_string(template, context)
    msg = EmailMessage(subject, message, to=to, from_email=from_email)
    msg.content_subtype = 'html'
    msg.send()
    return True


def test_data():
    for tr in Transaction.objects.all():
        tr.id = None
        tr.save()


@shared_task
def add(x, y):
    return x + y

def is_arc(id):
    country = Country.objects.get(id=id)
    if country.name == 'United States':
        return True
    return False



def country_import():
    countries = [
        {'name': 'Canada', 'code': 'CA', 'states': [
            ('Ontario', 'ON'),
            ('Quebec', 'QC'),
            ('Nova Scotia', 'NS'),
            ('New Brunswick', 'NB'),
            ('Manitoba', 'MB'),
            ('British Columbia', 'BC'),
            ('Prince Edward Island', 'PE'),
            ('Saskatchewan', 'SK'),
            ('Alberta', 'AB'),
            ('Newfoundland and Labrador', 'NL'),
            ('Yukon', 'YT'),
            ('Northwest Territories', 'NT'),
            ('Nunavut', 'NU'),
            ('Canada', 'CA'),
        ]},
        {'name': 'United States', 'code': 'US', 'states': [
            ('Alabama', 'AL'),
            ('Alaska', 'AK'),
            ('Arizona', 'AZ'),
            ('Arkansas', 'AR'),
            ('California', 'CA'),
            ('Colorado', 'CO'),
            ('Connecticut', 'CT'),
            ('Delaware', 'DE'),
            ('Florida', 'FL'),
            ('Georgia', 'GA'),
            ('Hawaii', 'HI'),
            ('Idaho', 'ID'),
            ('Illinois', 'IL'),
            ('Indiana', 'IN'),
            ('Iowa', 'IA'),
            ('Kansas', 'KS'),
            ('Kentucky', 'KY'),
            ('Louisiana', 'LA'),
            ('Maine', 'ME'),
            ('Maryland ', 'MD'),
            ('Massachusett', 'MA'),
            ('Michigan', 'MI'),
            ('Minnesota', 'MN'),
            ('Mississippi', 'MS'),
            ('Missouri', 'MO'),
            ('Montana', 'MT'),
            ('Nebraska', 'NE'),
            ('Nevada', 'NV'),
            ('New Hampshire', 'NH'),
            ('New Jersey', 'NJ'),
            ('New Mexico', 'NM'),
            ('New York', 'NY'),
            ('North Carolina', 'NC'),
            ('North Dakota', 'ND'),
            ('Ohio', 'OH'),
            ('Oklahoma', 'OK'),
            ('Oregon', 'OR'),
            ('Pennsylvania', 'PA'),
            ('Rhode Island', 'RI'),
            ('South Carolina', 'SC'),
            ('South Dakota', 'SD'),
            ('Tennessee', 'TN'),
            ('Texas', 'TX'),
            ('Utah', 'UT'),
            ('Vermont', 'VT'),
            ('Virginia', 'VA'),
            ('Washington', 'WA'),
            ('West Virginia', 'WV'),
            ('Wisconsin', 'WI'),
            ('Wyoming', 'WY'),
            ('Puerto Rico', 'PR'),
            ('Virgin Islands', 'VI'),
            ('Guam', 'GU'),
            ('Northern Mariana Islands', 'MP'),
            ('American Samoa', 'AS'),
            ('District of Columbia', 'DC'),
            ('Federated States of Micronesia', 'FM'),
            ('Marshall Islands', 'MH'),
            ('Palau', 'PW'),
            ('Armed Forces Africa', 'AE'),
            ('Armed Forces Americas', 'AA'),
            ('Armed Forces Canada', 'AE'),
            ('Armed Forces Europe', 'AE'),
            ('Armed Forces Middle East', 'AE'),
            ('Armed Forces Pacific', 'AP'),
            ('American Samoa', 'AS'),
            ('District of Columbia', 'DC'),
            ('Federated States of Micronesia', 'FM'),
            ('Marshall Islands', 'MH'),
            ('Palau', 'PW'),
            ('Armed Forces Africa', 'AE'),
            ('Armed Forces Americas', 'AA'),
            ('Armed Forces Canada', 'AE'),
            ('Armed Forces Europe', 'AE'),
            ('Armed Forces Middle East', 'AE'),
            ('Armed Forces Pacific', 'AP'),
        ]},
    ]

    for country in countries:
        c, created = Country.objects.get_or_create(
                name=country.get('name'),
                defaults={
                    'code': country.get('code'),
                }
            )
        for state in country.get('states'):
            st, created = State.objects.get_or_create(
                name=state[0],
                defaults={
                    'abrev': state[1],
                    'country': c
                }
            )