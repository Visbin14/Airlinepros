# Generated by Django 2.1.2 on 2019-01-21 13:29

from django.db import migrations


class Migration(migrations.Migration):

    dependencies = [
        ('report', '0019_auto_20190114_1231'),
    ]

    operations = [
        migrations.AlterModelOptions(
            name='reportfile',
            options={'permissions': (('view_sales_details', 'Can view sales details report'), ('download_sales_details', 'Can download sales details report'))},
        ),
    ]
