# Generated by Django 2.1.2 on 2019-03-11 07:00

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('report', '0036_deduction'),
    ]

    operations = [
        migrations.AlterModelOptions(
            name='reportfile',
            options={'permissions': (('view_sales_details', 'Can view sales details report'), ('download_sales_details', 'Can download sales details report'), ('view_sales_summary', 'Can view sales summary report'), ('download_sales_summary', 'Can download sales summary report'), ('view_adm', 'Can view adm report'), ('download_adm', 'Can download adm report'), ('view_sales_by', 'Can view sales by report'), ('download_sales_by', 'Can download sales by report'), ('view_all_sales', 'Can view all sales report'), ('download_all_sales', 'Can download all sales report'), ('view_year_to_year', 'Can view year to year report'), ('download_year_to_years', 'Can download year to year report'), ('view_commission', 'Can view commission report'), ('download_commission', 'Can download commission report'), ('view_sales_comparison', 'Can view sales comparison report'), ('download_sales_comparison', 'Can download sales comparison report'), ('view_top_agency', 'Can view top agency report'), ('download_top_agency', 'Can download top agency report'), ('view_monthly_yoy', 'Can view monthly yoy report'), ('download_monthly_yoy', 'Can download monthly yoy report'), ('view_airline_agency', 'Can view airline agency report'), ('download_airline_agency', 'Can download airline agency report'), ('view_agency_collection_report', 'Can view agency collection report'), ('download_agency_collection_report', 'Can download agency collection report'), ('view_upload_reports', 'Can upload report files'), ('view_upload_calendar', 'Can upload calendar files'), ('view_calendar', 'Can view calendar data'), ('view_disbursement_summary', 'Can view disbursement summary report'), ('download_disbursement_summary', 'Can download disbursement summary report'))},
        ),
        migrations.AddField(
            model_name='disbursement',
            name='arc_net_disb',
            field=models.FloatField(default=0.0),
        ),
    ]
