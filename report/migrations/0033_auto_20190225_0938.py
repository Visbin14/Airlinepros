# Generated by Django 2.1.2 on 2019-02-25 09:38

from django.db import migrations, models
import django.db.models.deletion


class Migration(migrations.Migration):

    dependencies = [
        ('main', '0027_auto_20190211_0703'),
        ('report', '0032_auto_20190222_0626'),
    ]

    operations = [
        migrations.CreateModel(
            name='Disbursement',
            fields=[
                ('id', models.AutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('filedate', models.DateField()),
                ('rundate1', models.DateField()),
                ('rundate2', models.DateField(null=True)),
                ('file1', models.FileField(upload_to='disbursements')),
                ('file2', models.FileField(null=True, upload_to='disbursements')),
                ('arc_deduction', models.FloatField()),
                ('arc_fees', models.FloatField()),
                ('arc_tot', models.FloatField()),
                ('arc_reversal', models.FloatField()),
                ('imported_at', models.DateTimeField(auto_now_add=True)),
                ('pending_deductions', models.BooleanField(default=False)),
                ('airline', models.ForeignKey(null=True, on_delete=django.db.models.deletion.CASCADE, to='main.Airline')),
                ('report_period', models.ForeignKey(null=True, on_delete=django.db.models.deletion.CASCADE, to='report.ReportPeriod')),
            ],
        ),
        migrations.AlterUniqueTogether(
            name='disbursement',
            unique_together={('airline', 'report_period'), ('airline', 'filedate')},
        ),
    ]
