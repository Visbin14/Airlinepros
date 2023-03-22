# Generated by Django 2.1.2 on 2019-02-25 09:39

from django.db import migrations, models
import django.db.models.deletion


class Migration(migrations.Migration):

    dependencies = [
        ('main', '0027_auto_20190211_0703'),
        ('report', '0033_auto_20190225_0938'),
    ]

    operations = [
        migrations.CreateModel(
            name='CarrierDeductions',
            fields=[
                ('id', models.AutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('filedate', models.DateField()),
                ('file', models.FileField(upload_to='deductions')),
                ('imported_at', models.DateTimeField(auto_now_add=True)),
                ('no_bill_items', models.IntegerField(null=True)),
                ('airline', models.ForeignKey(null=True, on_delete=django.db.models.deletion.CASCADE, to='main.Airline')),
                ('report_period', models.ForeignKey(null=True, on_delete=django.db.models.deletion.CASCADE, to='report.ReportPeriod')),
            ],
        ),
        migrations.AlterUniqueTogether(
            name='carrierdeductions',
            unique_together={('airline', 'report_period')},
        ),
    ]