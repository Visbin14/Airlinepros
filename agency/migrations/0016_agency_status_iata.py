# Generated by Django 2.1.2 on 2018-11-14 06:37

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('agency', '0015_agencylistreference_file_type'),
    ]

    operations = [
        migrations.AddField(
            model_name='agency',
            name='status_iata',
            field=models.CharField(blank=True, choices=[('A', 'Active'), ('D', 'Default Information'), ('R', 'Reviews/Notices of Termination'), ('S', 'Reinstatements'), ('T', 'Terminations And Closures')], default='A', max_length=1, null=True),
        ),
    ]
