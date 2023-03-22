from __future__ import absolute_import, unicode_literals
import os
from celery import Celery
from decouple import config

# set the default Django settings module for the 'celery' program.
from asplinks import settings
from celery.schedules import crontab
from datetime import timedelta, datetime, date

os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'asplinks.settings')

app = Celery('asplinks', broker=config('BROKER'))

# Using a string here means the worker don't have to serialize
# the configuration object to child processes.
# - namespace='CELERY' means all celery-related configuration keys
#   should have a `CELERY_` prefix.
app.config_from_object('django.conf:settings', namespace='CELERY')
from celery.task import task
# Load task modules from all registered Django app configs.
# app.autodiscover_tasks(lambda: settings.INSTALLED_APPS)
app.conf.update(
    result_backend='django-db',
)
app.conf.beat_schedule = {
    'task_download_files_scheduler': {
        'task': 'download_files_scheduler',
        'schedule': crontab(hour=11, minute=22),
    },
    'task_new_agent_report_generate':{
        'task': 'new_agent_report_generate',
        'schedule': crontab(hour=10, minute=30),
    }
}

# @app.task(bind=True)
# def debug_task(self):
#     print('Request: {0!r}'.format(self.request))


@task(name="download_files_scheduler")
def download_files(manual=False):
    from main.views import download_latest
    from main.models import RemoteServers,LatestFiles
    current_date = datetime.now()
    print(":::  ",current_date.weekday())
    flag = False
    if manual:
        flag = True
    elif not manual and current_date.weekday() == 3:
        flag = True
    if flag: # only on Wednesday
        # for ftp_obj in RemoteServers.objects.all():
        for ftp_obj in RemoteServers.objects.all(): 
            if LatestFiles.objects.filter(ftp_obj=ftp_obj):
                latestfile_obj = LatestFiles.objects.get(ftp_obj=ftp_obj)
                initial = False
                latest = latestfile_obj.latest
            else:
                initial = True
                latest = 0
            # try:
            download_latest(ftp_obj.hostname, ftp_obj.user, ftp_obj.password, ftp_obj.countrycode, target_path=None,
                            initial=initial, latest=latest, ftp_obj=ftp_obj)
            # except Exception as e:
            #     print("Error   :   ",e)



@task(name="new_agent_report_generate")
def new_agent_report_generate():
    from main.models import Airline
    from report.models import Transaction, Disbursement
    from agency.models import NewAgents
    day = date.today()
    one_day = timedelta(1)
    week_3_bef = datetime.now() - timedelta(days=21)
    NewAgents.objects.filter(ped__lte=week_3_bef.date()).delete()
    def week_number_of_month(date_value):
        return (date_value.isocalendar()[1] - date_value.replace(day=1).isocalendar()[1] + 1)
    
    while True:
        x = Disbursement.objects.filter(
            report_period__ped=day)
        if len(x):
            break
        day = day - one_day
    for airline in Airline.objects.filter(country__code='US'):
        transactions = Transaction.objects.filter(report__airline=airline)
        transaction_fltrd = transactions.filter(
            report__report_period__ped=day)
        for transaction in transaction_fltrd:
            other_trans = transactions.filter(
                agency=transaction.agency)
            if len(other_trans) == 1:
                week_num = week_number_of_month(day)
                NewAgents.objects.get_or_create(agency=transaction.agency,airline=airline,ped=day,week=week_num)