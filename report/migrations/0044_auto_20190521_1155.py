# Generated by Django 2.1.2 on 2019-05-21 11:55

from django.db import migrations, models
import django.db.models.deletion


class Migration(migrations.Migration):

    dependencies = [
        ('report', '0043_dailycreditcardfile_from_date'),
    ]

    operations = [
        migrations.AlterField(
            model_name='agencydebitmemo',
            name='transaction',
            field=models.OneToOneField(null=True, on_delete=django.db.models.deletion.CASCADE, to='report.Transaction'),
        ),
    ]