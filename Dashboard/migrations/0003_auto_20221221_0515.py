# Generated by Django 2.1.2 on 2022-12-21 05:15

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('Dashboard', '0002_auto_20221220_1647'),
    ]

    operations = [
        migrations.AlterField(
            model_name='dashboardmodel',
            name='Year',
            field=models.CharField(blank=True, max_length=50, null=True),
        ),
    ]
