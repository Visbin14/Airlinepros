# Generated by Django 2.1.2 on 2019-06-04 07:04

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('report', '0046_excelreportdownload'),
    ]

    operations = [
        migrations.AddField(
            model_name='transaction',
            name='is_sale',
            field=models.BooleanField(default=True),
        ),
    ]
