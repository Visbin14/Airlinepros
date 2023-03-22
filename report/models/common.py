from django.db import models
from main.models import Airline, BaseModel, Country
import calendar
import numpy as np

from math import ceil
from datetime import timedelta, date
from django.db.models.signals import post_delete
from django.dispatch import receiver
from agency.models import Agency

# calendar.setfirstweekday(6)


class ReportPeriod(models.Model):
    year = models.IntegerField()
    month = models.SmallIntegerField()
    week = models.SmallIntegerField()
    ped = models.DateField()
    from_date = models.DateField()
    remittance_date = models.DateField()
    country = models.ForeignKey(
        Country, null=True, blank=True, related_name='report_period', on_delete=models.SET_NULL)

    class Meta:
        unique_together = ('year', 'month', 'week', 'country', 'ped')
        indexes = [
            models.Index(fields=['ped', 'year', 'month', 'from_date', 'country']),
        ]

    def __str__(self):
        return "Week %s, %s > %s" % (self.week, date(self.year, self.month, 1).strftime("%b %Y"),self.country)

    # @classmethod
    # def from_date(self, report_date):
    #     """
    #     Returns a ReportPeriod object for the given date.
    #     """
    #     next_sun = next_sunday(report_date)
    #     try:
    #         return self.objects.get(ped=next_sun)
    #     except self.DoesNotExist:
    #         rp = self.objects.create(year=report_date.year,
    #                                  month=report_date.month,
    #                                  week=get_week_of_month(report_date.year, report_date.month, report_date.day),
    #                                  ped=next_sun)
    #         return rp


class ReportFile(models.Model):
    report_period = models.ForeignKey(
        ReportPeriod, on_delete=models.CASCADE, null=True)
    airline = models.ForeignKey(Airline, on_delete=models.CASCADE, null=True)
    file = models.FileField(upload_to='reportfile')

    ref_no = models.CharField(max_length=99, null=True)
    imported_at = models.DateTimeField(auto_now_add=True)

    transaction_amount = models.FloatField(null=True)
    fare_amount = models.FloatField(null=True)
    tax = models.FloatField(null=True)
    fandc = models.FloatField(null=True)
    pen = models.FloatField(null=True)
    cobl_amount = models.FloatField(null=True)
    std_comm = models.FloatField(null=True)
    supp_comm = models.FloatField(null=True)
    tax_on_comm = models.FloatField(null=True)
    balance = models.FloatField(null=True)
    country = models.ForeignKey(
        Country, null=True, blank=True, related_name='report_files', on_delete=models.SET_NULL)

    acms = models.FloatField(null=True, default=0.00)
    cc = models.FloatField(null=True, default=0.00)
    ca = models.FloatField(null=True, default=0.00)

    class Meta:
        # unique_together = (
        #     ('airline','report_period'),
        #     ('airline', 'filedate'),
        # )
        indexes = [
            models.Index(fields=['report_period', 'airline', 'transaction_amount', 'fare_amount', 'tax', 'fandc', 'pen',
                                 'cobl_amount', 'std_comm', 'supp_comm', 'tax_on_comm', 'balance', 'acms', 'cc', 'ca']),
        ]
        permissions = (
            ('view_sales_details', "Can view sales details report"),
            ('download_sales_details', "Can download sales details report"),
            ('view_sales_summary', "Can view sales summary report"),
            ('view_new_sales_summary',"Can view new sales summary report"),
            ('download_sales_summary', "Can download sales summary report"),
            ('download_new_sales_summary',"Can download new sales summary report"),
            ('view_adm', "Can view adm report"),
            ('download_adm', "Can download adm report"),
            ('view_sales_by', "Can view sales by report"),
            ('download_sales_by', "Can download sales by report"),
            ('view_all_sales', "Can view all sales report"),
            ('download_all_sales', "Can download all sales report"),
            ('view_year_to_year', "Can view year to year report"),
            ('download_year_to_years', "Can download year to year report"),
            ('view_commission', "Can view commission report"),
            ('download_commission', "Can download commission report"),
            ('view_sales_comparison', "Can view sales comparison report"),
            ('download_sales_comparison', "Can download sales comparison report"),
            ('view_top_agency', "Can view top agency report"),
            ('download_top_agency', "Can download top agency report"),
            ('view_monthly_yoy', "Can view monthly yoy report"),
            ('download_monthly_yoy', "Can download monthly yoy report"),
            ('view_airline_agency', "Can view airline agency report"),
            ('download_airline_agency', "Can download airline agency report"),
            ('view_agency_collection_report', "Can view agency collection report"),
            ('download_agency_collection_report', "Can download agency collection report"),
            ('view_upload_reports', "Can upload report files"),
            ('view_upload_calendar', "Can upload calendar files"),
            ('view_calendar', "Can view calendar data"),
            ('view_disbursement_summary', "Can view disbursement summary report"),
            ('download_disbursement_summary', "Can download disbursement summary report"),
            ('view_airline_management', "Can view airline management"),
            ('change_airline_management', "Can change airline management"),
            ('view_weekly_sales_report',"Can view weekly sales report"),
            ('download_weekly_sales',"Can download weekly sales report"),
        )

    def __str__(self):
        return "%s - %s" % (
            self.airline, self.ref_no)


class DailyCreditCardFile(BaseModel):
    airline = models.ForeignKey(Airline, on_delete=models.CASCADE, null=True)
    date = models.DateField()
    from_date = models.DateField(null=True, blank=True)
    grand_total = models.FloatField(null=True, default=0.00)

    def __str__(self):
        return "%s - %s" % (
            self.airline, self.date)


def week_of_month(dt):
    """
    Returns the week of the month for the specified date.
    """
    first_day = dt.replace(day=1)
    dom = dt.day
    adjusted_dom = dom + (1 + first_day.weekday()) % 7
    return int(ceil(adjusted_dom / 7.0))


def get_week_of_month(year, month, day):
    x = np.array(calendar.monthcalendar(year, month))
    week_of_monthh = np.where(x == day)[0][0] + 1
    return (week_of_monthh)


def next_sunday(dt):
    """
    Returns next sunday for the specified date.
    """
    return dt + timedelta(7 - dt.isoweekday())


@receiver(post_delete, sender=ReportFile)
def submission_delete(sender, instance, **kwargs):
    instance.file.delete(False)


class ReprocessFile(BaseModel):
    is_done = models.BooleanField(default=False)
    message = models.TextField(blank=True)
    status = models.CharField(blank=True, max_length=100)
    airline = models.ForeignKey(Airline, null=True, blank=True, on_delete=models.SET_NULL)
    start_date = models.DateField(blank=True, null=True)
    end_date = models.DateField(blank=True, null=True)

    def __str__(self):
        return '{} To {} of {}'.format(self.start_date, self.end_date, self.airline or 'All Airline')


class ExcelReportDownload(BaseModel):
    REPORT_CHOICES = [
        (1, 'Sales Details'),
        (2, 'Commission Report'),
        (3, 'ADM report'),
    ]

    file = models.FileField(upload_to='excelreports')
    report_type = models.IntegerField(choices=REPORT_CHOICES, default=1)

    def __str__(self):
        return str(self.id)
