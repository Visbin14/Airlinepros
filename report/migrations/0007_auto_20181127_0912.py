# Generated by Django 2.1.2 on 2018-11-27 09:12

from django.db import migrations


class Migration(migrations.Migration):

    dependencies = [
        ('report', '0006_auto_20181127_0859'),
    ]

    operations = [
        migrations.RenameField(
            model_name='reportfile',
            old_name='cobal_amount',
            new_name='cobl_amount',
        ),
    ]