import django
django.setup()
from asplinks.celery import download_files
download_files(manual=True)
