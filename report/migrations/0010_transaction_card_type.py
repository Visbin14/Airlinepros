# Generated by Django 2.1.2 on 2018-12-04 07:11

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('report', '0009_auto_20181128_1310'),
    ]

    operations = [
        migrations.AddField(
            model_name='transaction',
            name='card_type',
            field=models.CharField(blank=True, max_length=99, null=True),
        ),
    ]
