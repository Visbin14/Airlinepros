# python imports
import calendar
import datetime
import os
import tempfile
import re
import threading
import time
from urllib import request
import zipfile
from collections import OrderedDict
from subprocess import call
from zipfile import ZipFile

# django imports
from django.conf import settings
from django.contrib import messages
from django.contrib.auth.mixins import PermissionRequiredMixin
from django.core.files.storage import FileSystemStorage
from django.db.models import CharField, Count, F, FloatField, OuterRef, Q, Subquery, Sum, Value as V
from django.db.models.functions import Cast, Coalesce
from agency.models import Agency,AgencyCollection
from django.http import HttpResponse, HttpResponseRedirect, JsonResponse
from django.shortcuts import redirect, render
from django.views import View
from django.views.generic import FormView, ListView, TemplateView
from datetime import timedelta
from hotfileloader import process_billing_details_from_hotfile
# 3rd-party imports
import dateutil.relativedelta as relativedelta
import dateutil.rrule as rrule
import xlwt
from celery import task
from dateutil.relativedelta import relativedelta as rd

# util imports
from account.models import User
from main.excelstyle import *
from main.models import Airline, City, CommissionHistory, Country, State
from main.tasks import is_arc, send_mail
from report.forms import ReportFileForm
from report.models import (
    AgencyDebitMemo, CarrierDeductions, Charges, DailyCreditCardFile, Deduction, Disbursement, ExcelReportDownload,
    Remittance, ReportFile, ReportPeriod, ReprocessFile, Taxes, Transaction)
from report.tasks import (
    process_billing_details, process_card_details, process_carrier_deductions, process_carrier_report,
    process_disbursement_advice, process_excelfile, re_process)


def format_value_excel(val):
    return round(val, 2) if val else 0.00


def format_value(val):
    return "{:,.2f}".format(val) if val else '0.00'


def value_check(val):
    return val if val else 0.00


class ReportUpload(PermissionRequiredMixin, FormView):
    """
    ReportUpload allows user to upload report files and then it passes through various processes to get all required data from the file.
    """
    permission_required = ('report.view_upload_reports',)
    template_name = 'upload-sales-file.html'
    form_class = ReportFileForm
    success_url = '/reports/upload/'
    errMsgnw = ''
    miss_match_errors = dict()

    def get_context_data(self, **kwargs):

        tt = Transaction.objects.all().aggregate(
            transaction_amount=Coalesce(Sum('transaction_amount'), V(0)),
            fare_amount=Coalesce(Sum('fare_amount'), V(0)))
        context = super(ReportUpload, self).get_context_data(**kwargs)
        context["errMsg"] = self.errMsgnw
        context['is_arc'] = is_arc(self.request.session.get('country'))
        if self.miss_match_errors:
            context['miss_match_errors'] = self.miss_match_errors
        return context

    def form_valid(self, form):
        import threading
        media_root = getattr(settings, 'MEDIA_ROOT')
        fs = FileSystemStorage(location=os.path.join(media_root, "reportfile/"))

        def thread_process(self, country, form, file, filepath, filenw=None, from_zip=False):
            file_r = self.request.FILES['file'].read()
            # file_name, extention = file.name.split('.')
            splits=file.name.split('.')
            extention= splits[-1]
            file_name= file.name
            media_root = getattr(settings, 'MEDIA_ROOT')
            error = None
            if re.search("^[0-9]+$", extention) and country.name != 'United States':
                text_file = os.path.join(media_root, "reportfile/" + file_name)


                error= process_billing_details_from_hotfile(text_file, self.request)


            elif from_zip:
                extractfile_extension = extention
                extractfile_name = file_name
                if extractfile_extension.lower() == 'pdf' and country.name != 'United States':

                    lf = tempfile.NamedTemporaryFile(dir=os.path.join(media_root, "reportfile/"), suffix='.pdf')
                    lf.write(filenw.read(name))
                    text_file = os.path.join(media_root, "reportfile/" + extractfile_name + ".txt")
                    if str(extractfile_name).startswith(country.code + '_FCAIBILLDET'):
                        call(['pdftotext', '-layout', lf.name, text_file])
                        error = process_billing_details(text_file, self.request)
                    elif str(extractfile_name).startswith(country.code + '_PCAIDLYDET'):
                        call(['pdftotext', '-layout', lf.name, text_file])
                        error = process_card_details(text_file, self.request)
                    else:
                        error = "Incorrect file"
                    if error:
                        context = {
                            'request': self.request,
                            'error': error,
                            'file_name': name
                        }
                        admin_emails = list(User.objects.filter(is_superuser=True).values_list('email', flat=True))
                        # send_mail("File upload parsing issue.", "email/parsing-issue-email.html", context, admin_emails,
                        #           from_email='Assda@assda.com')
                    lf.close()


                elif extractfile_extension.lower() == 'txt' and country.name == 'United States':


                    filenw.extract(name, os.path.join(media_root, "reportfile/"))
                    file_name = extractfile_name

                    oldname = os.path.join(os.path.join(media_root, "reportfile/"), name)
                    stored_filename = file_name + datetime.datetime.now().strftime('%s') + '.txt'

                    filepath = os.path.join(os.path.join(media_root, "reportfile/"), stored_filename)

                    os.rename(oldname, filepath)

                    if str(file_name).startswith('CARRPTSW'):
                        error = process_carrier_report(filepath, self.request)
                    elif str(file_name).startswith('DISBADV'):
                        error = process_disbursement_advice(filepath, file_name, self.request)
                    elif str(file_name).startswith('CARRDED'):
                        error = process_carrier_deductions(filepath, file_name, self.request)
                    else:
                        error = "Incorrect file"

                    if error:
                        if os.path.exists(filepath):
                            os.remove(filepath)
                        context = {
                            'request': self.request,
                            'error': error,
                            'file_name': name
                        }
                        admin_emails = list(User.objects.filter(is_superuser=True).values_list('email', flat=True))
                        # send_mail("File upload parsing issue.", "email/parsing-issue-email.html", context, admin_emails,
                        #           from_email='Assda@assda.com')


            else:
                if extention.lower() == 'pdf' and country.name != 'United States':
                    # pdf file for IATA
                    text_file = os.path.join(media_root, "reportfile/" + file_name + ".txt")
                    lf = tempfile.NamedTemporaryFile(dir=os.path.join(media_root, "reportfile/"), suffix='.pdf')
                    lf.write(file_r)


                    if str(file_name).startswith(country.code + '_FCAIBILLDET'):
                        call(['pdftotext', '-layout', lf.name, text_file])
                        with open(text_file, 'r+', encoding="utf-8") as fd:
                            lines = fd.readlines()
                            fd.seek(0)
                            fd.writelines(line for line in lines if line.strip())
                            fd.truncate()
                        error = process_billing_details(text_file, self.request)
                    elif str(file_name).startswith(country.code + '_PCAIDLYDET'):
                        call(['pdftotext', '-layout', lf.name, text_file])
                        error = process_card_details(text_file, self.request)
                    else:
                        error = "Incorrect file"
                    # remove tmp file
                    lf.close()
                    if error:
                        if from_zip:
                            return False
                        else:
                            if not isinstance(error, dict):
                                form.add_error('file', error)
                                messages.add_message(self.request, messages.ERROR, error)
                            else:
                                self.miss_match_errors = error
                                messages.add_message(self.request, messages.ERROR, 'Mismatch in parsed amounts.')
                                context = {
                                    'request': self.request,
                                    'error': error,
                                    'file_name': file.name
                                }
                                admin_emails = list(
                                    User.objects.filter(is_superuser=True).values_list('email', flat=True))
                                # send_mail("File upload parsing issue.", "email/parsing-issue-email.html", context, admin_emails,
                                #           from_email='Assda@assda.com')


                            return super(ReportUpload, self).form_invalid(form)
                    messages.add_message(self.request, messages.SUCCESS, 'File uploaded successfully.')

                    return super(ReportUpload, self).form_valid(form)

                elif extention.lower() == 'txt' and country.name == 'United States':


                    if str(file_name).startswith('CARRPTSW'):
                        error = process_carrier_report(filepath, self.request)
                    elif str(file_name).startswith('DISBADV'):
                        error = process_disbursement_advice(filepath, file_name, self.request)
                    elif str(file_name).startswith('CARRDED'):
                        error = process_carrier_deductions(filepath, file_name, self.request)
                    else:
                        error = "Incorrect file"
                    # ARC

                    if error:
                        form.add_error('file', error)
                        messages.add_message(self.request, messages.ERROR, error)
                        return super(ReportUpload, self).form_invalid(form)
                    messages.add_message(self.request, messages.SUCCESS, 'File uploaded successfully.')
                    return super(ReportUpload, self).form_valid(form)

                else:
                    form.add_error('file', 'Unsupported file format.')
                    if country.name != 'United States':
                        messages.add_message(self.request, messages.ERROR,
                                             'Unsupported file format. Please upload zip file/pdf document.')
                        # send_mail("Unsupported file format. Please upload zip file/pdf document.",
                        #           "email/parsing-issue-email.html", {},
                        #           [self.request.user.email],
                        #           from_email='Assda@assda.com')
                    else:
                        messages.add_message(self.request, messages.ERROR,
                                             'Unsupported file format. Please upload zip file/txt document.')
                        # send_mail("Unsupported file format. Please upload zip file/txt document.", "email/parsing-issue-email.html", {},
                        #           [self.request.user.email],
                        #           from_email='Assda@assda.com')
                    return super(ReportUpload, self).form_invalid(form)

        file = form.files.get("file")
        #file_name, extention = file.name.split('.')
        splits=file.name.split('.')
        extention= splits[-1]
        file_name= file.name
        if self.request.POST.get("from_scheduler") is not None:
            country = Country.objects.get(id=self.request.POST.get("countrycode"))
        else:
            country = Country.objects.get(id=self.request.session.get('country'))
        if extention.lower() == 'zip':
            errfiles = []
            errMsg = ''
            zipfileCnt = 0

            dt = datetime.datetime.now()
            timestamp = int(time.mktime(dt.timetuple()))

            zipFldr = file_name + '_' + str(timestamp)
            filename = fs.save(zipFldr + '.' + extention, file)

            filenw = ZipFile(os.path.join(media_root, "reportfile/") + filename, 'r')
            cc = filenw.infolist()

            for name in filenw.namelist():

                if name.endswith('/'):
                    pass
                else:
                    zipfileCnt = len(filenw.namelist())
                    file_data = filenw.open(name)
                    from_zip = True
                    filepath = ""
                    process = threading.Thread(target=thread_process,
                                               args=(self, country, form, file_data, filepath, filenw, from_zip))
                    process.start()
                    process.join()


            if errfiles:
                temp = {'response': errfiles, 'request': self}


            elif zipfileCnt == 0:
                errMsg = 'No files found'
                form.add_error('file', errMsg)
                messages.add_message(self.request, messages.ERROR, errMsg)

                return super(ReportUpload, self).form_invalid(form)


        else:
            filepath = ""

            if re.search("^[0-9]+$", extention) and country.name != 'United States':
                stored_filename = file_name
                filepath = os.path.join(os.path.join(media_root, "reportfile/"), file_name)
                filesave = open(filepath, 'wb+')
                for i in file:
                    filesave.write(i)
                filesave.close()
            elif extention.lower() == 'txt' and country.name == 'United States':
                stored_filename = file_name + datetime.datetime.now().strftime('%s') + '.txt'
                filepath = os.path.join(os.path.join(media_root, "reportfile/"), stored_filename)
                filesave = open(filepath, 'wb+')
                for chunk in file.chunks():
                    filesave.write(chunk)
                filesave.close()
            process = threading.Thread(target=thread_process, args=(self, country, form, file, filepath))
            process.start()

        messages.add_message(self.request, messages.SUCCESS,
                             'This may take few minutes. We will send you an email once the upload is finished.')
        return super(ReportUpload, self).form_valid(form)




class SalesReport(PermissionRequiredMixin, ListView):
    """sales details listing with pagination."""

    permission_required = ('report.view_sales_details',)
    model = Transaction
    template_name = 'sales-report.html'
    context_object_name = 'transactions'
    paginate_by = 200
    all_peds = []
    all_days = []
    peds = []
    credit_file_dates = []
    disbursement_files = []
    no_peds = False





    def get_queryset(self):
        month_year = self.request.GET.get('month_year', '')
        start_date = self.request.GET.get('start_date', '')
        end_date = self.request.GET.get('end_date', '')
        airline = self.request.GET.get('airline', '')
        qs = Transaction.objects.select_related('agency').exclude(
            transaction_type__startswith='ACM', agency__agency_no__in=['6998051', ])



        country = self.request.session.get('country')
        is_arc_var = is_arc(country)
        self.is_arc_var = is_arc_var

        if airline:
            qs = qs.filter(report__airline=airline)



        if month_year:

            month = datetime.datetime.strptime(month_year, '%B %Y').month or ''
            year = datetime.datetime.strptime(month_year, '%B %Y').year or ''

            if month and year:
                qs = qs.filter(report__report_period__month=month, report__report_period__year=year)
                #qs= qs.filter(issue_date__year= year).filter(issue_date__month= month)


            end_day = calendar.monthrange(year, month)[1]
            before = datetime.datetime(year, month, 1)
            after = datetime.datetime(year, month, end_day)

            self.all_peds = list(ReportPeriod.objects.filter(year=year, month=month, country=country).values_list('ped',
                                                                                                                  flat=True).order_by(
                'ped'))
            if self.all_peds:
                first_ped_start = self.all_peds[0] + datetime.timedelta(days=-6)
                self.all_days = [(first_ped_start + datetime.timedelta(days=x)) for x in
                                 range(1, (self.all_peds[-1] - first_ped_start).days + 2)]

                if not self.is_arc_var:
                    self.credit_file_dates = [cf.date for cf in
                                              DailyCreditCardFile.objects.filter(airline=airline,
                                                                                 date__range=[self.all_days[0],
                                                                                              self.all_days[
                                                                                                  -1]])]
                    self.disbursement_files = []
                else:
                    self.credit_file_dates = list(CarrierDeductions.objects.filter(airline=airline,
                                                                                   filedate__range=[self.all_days[0],
                                                                                                    self.all_days[
                                                                                                        -1]]).values_list(
                        'filedate', flat=True))
                    self.disbursement_files = list(Disbursement.objects.filter(airline=airline,
                                                                               filedate__range=[self.all_days[0],
                                                                                                self.all_days[
                                                                                                    -1]]).values_list(
                        'filedate', flat=True))

            else:
                self.no_peds = True

        if start_date and end_date:
            start = datetime.datetime.strptime(start_date, '%d %B %Y')
            end = datetime.datetime.strptime(end_date, '%d %B %Y')

            qs = qs.filter(report__report_period__ped__range=[start, end])

            self.all_peds = list(
                ReportPeriod.objects.filter(ped__range=[start, end]).values_list('ped', flat=True).order_by('ped'))
            if self.all_peds:
                first_ped_start = self.all_peds[0] + datetime.timedelta(days=-6)
                self.all_days = [(first_ped_start + datetime.timedelta(days=x)) for x in
                                 range(1, (self.all_peds[-1] - first_ped_start).days + 2)]
                if not self.is_arc_var:
                    self.credit_file_dates = [cf.date for cf in DailyCreditCardFile.objects.filter(airline=airline,
                                                                                                   date__range=[
                                                                                                       self.all_days[0],
                                                                                                       self.all_days[
                                                                                                           -1]])]
                    self.disbursement_files = []
                else:
                    self.credit_file_dates = list(CarrierDeductions.objects.filter(airline=airline,
                                                                                   filedate__range=[self.all_days[0],
                                                                                                    self.all_days[
                                                                                                        -1]]).values_list(
                        'filedate', flat=True))
                    self.disbursement_files = list(Disbursement.objects.filter(airline=airline,
                                                                               filedate__range=[self.all_days[0],
                                                                                                self.all_days[
                                                                                                    -1]]).values_list(
                        'filedate', flat=True))
            else:
                self.no_peds = True
        self.peds = list(qs.values_list('report__report_period__ped', flat=True).distinct())
        if self.is_arc_var:
            self.all_days = self.all_peds

        return qs



    def get_cancel_pen(self):
        month_year = self.request.GET.get('month_year', '')
        start_date = self.request.GET.get('start_date', '')
        end_date = self.request.GET.get('end_date', '')
        airline = self.request.GET.get('airline', '')

        qs = Transaction.objects.select_related('agency').prefetch_related('taxes_set').exclude(
            transaction_type__startswith='ACM',
            agency__agency_no__in=['6998051', ]).annotate(
            yq=Sum('charges__amount', filter=Q(charges__type='YQ')),
            cp=Sum('charges__amount', filter=Q(charges__type='CP')),
            yr=Sum('charges__amount', filter=Q(charges__type='YR')))

        if airline:


            qs = qs.filter(report__airline=airline)


        if month_year:

            month = datetime.datetime.strptime(month_year, '%B %Y').month or ''
            year = datetime.datetime.strptime(month_year, '%B %Y').year or ''

            if month and year:


                qs = qs.filter(report__report_period__month=month, report__report_period__year=year)
                #qs= qs.filter(issue_date__year= year).filter(issue_date__month= month).order_by('issue_date')



        if start_date and end_date:
            start = datetime.datetime.strptime(start_date, '%d %B %Y')
            end = datetime.datetime.strptime(end_date, '%d %B %Y')

            qs = qs.filter(report__report_period__ped__range=[start, end])
            #qs = qs.filter(issue_date__range=[start, end]).order_by('issue_date')
            # for q in qs:
            #     print("_____________________________q-",q)

        return qs



    def get_context_data(self, **kwargs):
        if self.no_peds:
            messages.add_message(self.request, messages.WARNING, 'No PEDs in this period.')
        context = super(SalesReport, self).get_context_data(**kwargs)
        context['activate'] = 'reports'
        context['all_peds'] = self.all_peds
        context['all_days'] = self.all_days
        context['credit_file_dates'] = self.credit_file_dates
        context['disbursement_files_dates'] = self.disbursement_files
        context['peds'] = self.peds
        context['missing_peds'] = [date for date in self.all_peds if date not in self.peds]
        context['missing_creds'] = [date for date in self.all_days if date not in self.credit_file_dates]
        if self.is_arc_var:
            try:
                one_day = self.all_days[0]
                context['missing_creds'] = [one_day] if one_day not in self.credit_file_dates else []
            except Exception as e:
                pass
        if self.is_arc_var:
            context['missing_disb'] = [date for date in self.all_days if date not in self.disbursement_files]
        else:
            context['missing_disb'] = []
        context['missing_count'] = len(context['missing_creds']) + len(context['missing_peds']) + len(
            context['missing_disb'])
        context['order_by'] = self.request.GET.get('order_by', 'id')
        context['query'] = self.request.GET.get('q', '')
        context['sales_version'] = self.request.GET.get('sales_version', '')
        context['month_year'] = self.request.GET.get('month_year', '')
        context['selected_airline'] = self.request.GET.get('airline', '')
        context['date_filter'] = self.request.GET.get('date_filter', 'month_year')
        context['start_date'] = self.request.GET.get('start_date', '')
        context['end_date'] = self.request.GET.get('end_date', '')
        context['airlines'] = Airline.objects.filter(country=self.request.session.get('country'))
        cp=self.get_cancel_pen()

        context['new_transactions']=cp
        context['transaction_types_count'] = self.get_queryset().aggregate(
            tickets=Count('pk', filter=Q(transaction_type='TKTT') | Q(transaction_type='TKT')),
            refunds=Count('pk', filter=Q(transaction_type='RFND')),
            exchanges=Count('pk', filter=Q(fop='EX') | Q(transaction_type='EXCH')))

        return context



class SalesByReport(PermissionRequiredMixin, ListView):
    """Sales by report."""

    permission_required = ('report.view_sales_by',)
    model = Transaction
    template_name = 'sales-by-report.html'
    context_object_name = 'transactions'

    def get_queryset(self):
        state = self.request.GET.get('state', '')
        organize_by = self.request.GET.get('organize_by', '')
        start_date = self.request.GET.get('start_date', '')
        end_date = self.request.GET.get('end_date', '')
        airline = self.request.GET.get('airline', '')
        qs = Transaction.objects.select_related('agency', 'report').filter(is_sale=True)

        if airline:
            qs = qs.filter(report__airline=airline)
        else:
            qs = None
        if state:
            qs = qs.filter(agency__state=state)

        if start_date and end_date:
            start = datetime.datetime.strptime(start_date, '%d %B %Y')
            end = datetime.datetime.strptime(end_date, '%d %B %Y')
            qs = qs.filter(report__report_period__ped__range=[start, end])

            if organize_by == 'agency':
                qs = qs.values('agency').order_by().distinct().annotate(total=Sum('fare_amount'),
                                                                        total_pen=Sum('pen',
                                                                                      filter=Q(pen_type='CANCEL PEN')),
                                                                        agency_trade_name=F('agency__trade_name'),
                                                                        agency_no=F('agency__agency_no'),
                                                                        sales_owner=F('agency__sales_owner__email'),
                                                                        state=F('agency__state__name'),
                                                                        tel=F('agency__tel'),
                                                                        agency_type=F('agency__agency_type__name'))

            if organize_by == 'state':
                qs = qs.values('agency__state').order_by().distinct().annotate(total=Sum('fare_amount'),
                                                                               total_pen=Sum('pen', filter=Q(
                                                                                   pen_type='CANCEL PEN')),
                                                                               sales_owner=F(
                                                                                   'agency__state__owner__email'),
                                                                               state=F('agency__state__name'))

            if organize_by == 'city':
                qs = qs.values('agency__city').order_by().distinct().annotate(total=Sum('fare_amount'),
                                                                              total_pen=Sum('pen', filter=Q(
                                                                                  pen_type='CANCEL PEN')),
                                                                              state_owner=F(
                                                                                  'agency__state__owner__email'),
                                                                              city=F('agency__city__name'),
                                                                              state_abrev=F('agency__state__abrev'))

            if organize_by == 'sales owner':
                qs = qs.values('agency__sales_owner').order_by().distinct().annotate(total=Sum('fare_amount'),
                                                                                     total_pen=Sum('pen', filter=Q(
                                                                                         pen_type='CANCEL PEN')),
                                                                                     sales_owner=F(
                                                                                         'agency__sales_owner__email'))

            if organize_by == 'agency_type':
                qs = qs.values('agency__agency_type').order_by().distinct().annotate(total=Sum('fare_amount'),
                                                                                     total_pen=Sum('pen', filter=Q(
                                                                                         pen_type='CANCEL PEN')),
                                                                                     agency_type=F(
                                                                                         'agency__agency_type__name'))


        return qs

    def get_context_data(self, **kwargs):
        context = super(SalesByReport, self).get_context_data(**kwargs)
        context['activate'] = 'reports'
        if self.request.GET.get('airline', ''):
            context['grand_total'] = self.get_queryset().aggregate(Sum('total'))
            context['grand_total_pen'] = self.get_queryset().aggregate(Sum('total_pen'))

        context['order_by'] = self.request.GET.get('order_by', 'id')
        context['selected_airline'] = self.request.GET.get('airline', '')
        context['selected_state'] = self.request.GET.get('state', '')
        context['organize_by'] = self.request.GET.get('organize_by', 'agency')
        context['start_date'] = self.request.GET.get('start_date', '')
        context['end_date'] = self.request.GET.get('end_date', '')
        context['airlines'] = Airline.objects.filter(country=self.request.session.get('country')).order_by('name')
        context['states'] = State.objects.filter(country=self.request.session.get('country')).order_by('name')
        context['is_arc'] = is_arc(self.request.session.get('country'))
        return context


class AllSalesReport(PermissionRequiredMixin, View):
    """All Sales report."""

    permission_required = ('report.view_all_sales',)
    model = Transaction
    template_name = 'all-sales-report.html'
    context_object_name = 'transactions'

    def get(self, request):
        month_year = self.request.GET.get('month_year', '')
        start_date = self.request.GET.get('start_date', '')
        end_date = self.request.GET.get('end_date', '')
        sales_type = self.request.GET.get('sales_type', '')
        is_arc_var = is_arc(self.request.session.get('country'))

        if is_arc_var and sales_type == 'Net':
            qs = Transaction.objects.select_related('agency').filter(~Q(transaction_type__in=['SP', 'ACM', 'ADM']),
                                                                     ~Q(agency__agency_no='6999001'))
        else:
            qs = Transaction.objects.select_related('agency').exclude(transaction_type__startswith='ACM',
                                                                      agency__agency_no='6998051')

        qs = qs.filter(is_sale=True)

        airline = self.request.GET.get('airline', 0)
        if airline:
            qs = qs.filter(report__airline=airline)

        if start_date and end_date:
            start = datetime.datetime.strptime(start_date, '%d %B %Y')
            end = datetime.datetime.strptime(end_date, '%d %B %Y')
            qs = qs.filter(report__report_period__ped__range=[start, end])

        if month_year:

            month = datetime.datetime.strptime(month_year, '%B %Y').month or ''
            year = datetime.datetime.strptime(month_year, '%B %Y').year or ''
            if month and year:
                qs = qs.filter(report__report_period__month=month, report__report_period__year=year)

        context = dict()
        context['peds'] = qs.values_list('report__report_period__ped', flat=True).distinct().order_by(
            'report__report_period__ped')
        context['airlines_values'] = qs.values('report__airline').annotate(name=F('report__airline__name'),
                                                                           str_id=Cast(F('report__airline__id'),
                                                                                       CharField())).distinct().order_by(
            'report__airline')
        if sales_type == 'Gross':
            annotations = dict()
            for airline in context['airlines_values']:
                annotation_name = '{}'.format(airline.get('report__airline'))
                annotations[annotation_name] = Sum('transaction_amount',
                                                   filter=Q(report__airline=airline.get('report__airline'))) - Sum(
                    'std_comm_amount', filter=Q(
                        report__airline=airline.get('report__airline')))
            valss = qs.values('report__report_period__ped').order_by('report__report_period__ped').distinct().annotate(
                **annotations)
            context['totals'] = qs.values('report__airline').annotate(
                total=Sum('transaction_amount') - Sum('std_comm_amount')).order_by('report__airline')
            context['value_list'] = valss
        else:
            if not is_arc_var:
                sub_tax_yq_yr = Charges.objects.select_related('transaction').filter(type__in=['YQ', 'YR'],
                                                                                     transaction__report__report_period__ped=OuterRef(
                                                                                         'report__report_period__ped'),
                                                                                     transaction__report__airline=airline,
                                                                                     transaction__is_sale=True).values(
                    'transaction__report__report_period__ped').annotate(
                    dcount=Count('transaction__report__report_period__ped'),
                    total_tax_yq_yr=Sum('amount', output_field=FloatField())).values('total_tax_yq_yr')
                sub_tax = Taxes.objects.select_related('transaction').filter(
                    transaction__report__report_period__ped=OuterRef('report__report_period__ped'),
                    transaction__report__airline=airline, transaction__is_sale=True).values(
                    'transaction__report__report_period__ped').annotate(
                    dcount=Count('transaction__report__report_period__ped'),
                    total_tax=Sum('amount', output_field=FloatField())).values('total_tax')

                context['vals'] = qs.values('report__airline', 'report__report_period__ped').order_by('report__airline',
                                                                                                      'report__report_period__ped').annotate(
                    airline_name=F('report__airline__name'),
                    net_sale=Sum('fare_amount') - Sum('std_comm_amount') + Coalesce(Sum('pen'), V(0)),
                    total=Sum('transaction_amount') - Sum('std_comm_amount'),
                    yqyr=Coalesce(Subquery(sub_tax_yq_yr, output_field=FloatField()), V(0)),
                    taxes=Coalesce(Subquery(sub_tax, output_field=FloatField()), V(0)))
                context['totals'] = context['vals'].aggregate(yqyr_total=Sum('yqyr'), net_total=Sum('net_sale'),
                                                              tax_total=Sum('taxes'), gross_total=Sum('total'))
            else:
                if airline:
                    sub_tax_yq_yr = Charges.objects.select_related('transaction').filter(type__in=['YQ', 'YR'],
                                                                                         transaction__report__report_period__ped=OuterRef(
                                                                                             'report__report_period__ped'),
                                                                                         transaction__report__airline=airline,
                                                                                         transaction__is_sale=True).values(
                        'transaction__report__report_period__ped').annotate(
                        dcount=Count('transaction__report__report_period__ped'),
                        total_tax_yq_yr=Sum('amount', output_field=FloatField())).values('total_tax_yq_yr')
                    sub_tax = Taxes.objects.select_related('transaction').filter(
                        transaction__report__report_period__ped=OuterRef('report__report_period__ped'),
                        transaction__report__airline=airline, transaction__is_sale=True).values(
                        'transaction__report__report_period__ped').annotate(
                        dcount=Count('transaction__report__report_period__ped'),
                        total_tax=Sum('amount', output_field=FloatField())).values('total_tax')

                    sub_acm = Transaction.objects.select_related('report', 'report__airline',
                                                                 'report__report_period').filter(
                        report__airline=airline, transaction_type__startswith='ACM', agency__agency_no='31768026',
                        report__report_period__ped=OuterRef('report__report_period__ped')).values(
                        'report__report_period__ped').annotate(
                        dcount=Count('report__report_period__ped'),
                        total_ap_acm=Sum('transaction_amount', output_field=FloatField())).values('total_ap_acm')
                    context['vals'] = qs.values('report__airline', 'report__report_period__ped').order_by(
                        'report__airline',
                        'report__report_period__ped').annotate(
                        airline_name=F('report__airline__name'),
                        net_sale=Coalesce(Sum('fare_amount'), V(0)) + Coalesce(Sum('pen'), V(0)),
                        yqyr=Coalesce(Subquery(sub_tax_yq_yr, output_field=FloatField()), V(0)),
                        taxes=Coalesce(Subquery(sub_tax, output_field=FloatField()), V(0)),
                        total=Sum('transaction_amount') - Sum('std_comm_amount') + Coalesce(
                            Subquery(sub_acm, output_field=FloatField()), V(0)))
                    context['totals'] = context['vals'].aggregate(yqyr_total=Sum('yqyr'), net_total=Sum('net_sale'),
                                                                  tax_total=Sum('taxes'), gross_total=Sum('total'))

        context['activate'] = 'reports'
        context['sales_type'] = self.request.GET.get('sales_type', '')
        context['month_year'] = self.request.GET.get('month_year', '')
        context['airlines'] = Airline.objects.filter(country=self.request.session.get('country'))
        context['selected_airline'] = self.request.GET.get('airline', '')
        context['date_filter'] = self.request.GET.get('date_filter', 'month_year')
        context['start_date'] = self.request.GET.get('start_date', '')
        context['end_date'] = self.request.GET.get('end_date', '')
        return render(request, self.template_name, context)


class MonthlyYOYReport(PermissionRequiredMixin, ListView):
    """Monthly YOY report ."""

    permission_required = ('report.view_monthly_yoy',)
    model = Transaction
    template_name = 'monthly-yoy-report.html'
    context_object_name = 'transactions'

    def get_queryset(self):
        organize_by = self.request.GET.get('organize_by', '')
        airline = self.request.GET.get('airline', '')
        is_arc_var = is_arc(self.request.session.get('country'))

        if is_arc_var and organize_by == 'net':
            qs = Transaction.objects.select_related('agency').filter(~Q(transaction_type__in=['SP', 'ACM', 'ADM']),
                                                                     ~Q(agency__agency_no='6999001'))
        else:
            qs = Transaction.objects.select_related('agency').exclude(transaction_type__startswith='ACM',
                                                                      agency__agency_no='6998051')

        qs = qs.filter(is_sale=True)

        years = self.request.GET.getlist('years')
        years.sort()
        months = [calendar.month_abbr[i] for i in range(1, 13)]

        annotations = {}
        annotations_net = {}
        result = []
        years_in_db = []

        if is_arc_var:
            for y in years:
                for i in range(1, 13):
                    annotation_name = '{}'.format(months[i - 1])
                    annotations[annotation_name] = Sum('transaction_amount',
                                                       filter=Q(report__report_period__month=i)) - Sum(
                        'std_comm_amount',
                        filter=Q(
                            report__report_period__month=i))

                    annotations_net[annotation_name] = Sum('fare_amount',
                                                           filter=Q(report__report_period__month=i)) + Sum('pen',
                                                                                                           filter=Q(
                                                                                                               report__report_period__month=i))
        else:
            for y in years:
                for i in range(1, 13):
                    annotation_name = '{}'.format(months[i - 1])
                    annotations[annotation_name] = Sum('transaction_amount',
                                                       filter=Q(report__report_period__month=i)) - Sum(
                        'std_comm_amount',
                        filter=Q(
                            report__report_period__month=i))

                    annotations_net[annotation_name] = Sum('fare_amount',
                                                           filter=Q(report__report_period__month=i)) - Sum(
                        'std_comm_amount', filter=Q(report__report_period__month=i)) + Sum('pen', filter=Q(
                        report__report_period__month=i))

        if airline:
            qs = qs.filter(report__airline=airline)

        if organize_by == 'gross':
            qs = qs.filter(report__report_period__year__in=years).values(
                'report__report_period__year').order_by().distinct().annotate(
                dcount=Count('report__report_period__year'), **annotations)

        if organize_by == 'net':
            qs = qs.filter(report__report_period__year__in=years).values(
                'report__report_period__year').order_by().distinct().annotate(
                dcount=Count('report__report_period__year'),
                **annotations_net)

        if organize_by:
            for obj in qs:
                od = OrderedDict()
                od['year'] = obj.get('report__report_period__year')
                years_in_db.append(str(obj.get('report__report_period__year')))
                for month in months:
                    od[month] = format_value(obj.get(month))
                result.append(od)

            for y in (set(years) - set(years_in_db)):
                od = OrderedDict()
                od['year'] = y
                for month in months:
                    od[month] = 0.00
                result.append(od)

        return sorted(result, key=lambda k: int(k['year']))

    def get_context_data(self, **kwargs):
        context = super(MonthlyYOYReport, self).get_context_data(**kwargs)
        context['activate'] = 'reports'
        context['q_string'] = self.request.META['QUERY_STRING']
        context['order_by'] = self.request.GET.get('order_by', 'id')
        context['selected_airline'] = self.request.GET.get('airline', '')
        context['organize_by'] = self.request.GET.get('organize_by', '')
        context['airlines'] = Airline.objects.filter(country=self.request.session.get('country'))
        context['years'] = list(range(2000, datetime.datetime.now().year + 1))
        context['months'] = [calendar.month_abbr[i] for i in range(1, 13)]
        years = self.request.GET.getlist('years')
        years.sort()
        context['selected_years'] = years
        return context


class GetMonthlyYOYReport(PermissionRequiredMixin, View):
    """Monthly YOY report ."""

    permission_required = ('report.download_monthly_yoy',)

    def get(self, request, *args, **kwargs):
        airline_name = ''
        response = HttpResponse(content_type='application/vnd.ms-excel')

        organize_by = self.request.GET.get('organize_by', '')
        airline = self.request.GET.get('airline', '')
        is_arc_var = is_arc(self.request.session.get('country'))

        if is_arc_var and organize_by == 'net':
            qs = Transaction.objects.select_related('agency').filter(~Q(transaction_type__in=['SP', 'ACM', 'ADM']),
                                                                     ~Q(agency__agency_no='6999001'))
        else:
            qs = Transaction.objects.select_related('agency').exclude(transaction_type__startswith='ACM',
                                                                      agency__agency_no='6998051')

        qs = qs.filter(is_sale=True)

        years = self.request.GET.getlist('years')
        years.sort()
        months = [calendar.month_abbr[i] for i in range(1, 13)]

        annotations = {}
        annotations_net = {}
        result = []
        years_in_db = []

        if is_arc_var:
            for y in years:
                for i in range(1, 13):
                    annotation_name = '{}'.format(months[i - 1])
                    annotations[annotation_name] = Sum('transaction_amount',
                                                       filter=Q(report__report_period__month=i)) - Sum(
                        'std_comm_amount',
                        filter=Q(
                            report__report_period__month=i))

                    annotations_net[annotation_name] = Sum('fare_amount',
                                                           filter=Q(report__report_period__month=i)) + Sum('pen',
                                                                                                           filter=Q(
                                                                                                               report__report_period__month=i))
        else:
            for y in years:
                for i in range(1, 13):
                    annotation_name = '{}'.format(months[i - 1])
                    annotations[annotation_name] = Sum('transaction_amount',
                                                       filter=Q(report__report_period__month=i)) - Sum(
                        'std_comm_amount',
                        filter=Q(
                            report__report_period__month=i))

                    annotations_net[annotation_name] = Sum('fare_amount',
                                                           filter=Q(report__report_period__month=i)) - Sum(
                        'std_comm_amount', filter=Q(report__report_period__month=i)) + Sum('pen', filter=Q(
                        report__report_period__month=i))
        if airline:
            qs = qs.filter(report__airline=airline)
            airline_data = Airline.objects.get(id=airline)
            airline_name = airline_data.name
            response[
                'Content-Disposition'] = 'inline; filename=' + airline_data.abrev + ' Monthly YOY report ' + '' + '.xls'

        if organize_by == 'gross':
            report_header = 'YOY COMPARISON OF GROSS SALES REPORT'
            qs = qs.filter(report__report_period__year__in=years).values(
                'report__report_period__year').order_by().distinct().annotate(
                dcount=Count('report__report_period__year'), **annotations)

        if organize_by == 'net':
            report_header = 'YOY COMPARISON OF NET SALES REPORT'
            qs = qs.filter(report__report_period__year__in=years).values(
                'report__report_period__year').order_by().distinct().annotate(
                dcount=Count('report__report_period__year'),
                **annotations_net)

        for obj in qs:
            od = OrderedDict()
            od['year'] = obj.get('report__report_period__year')
            years_in_db.append(str(obj.get('report__report_period__year')))
            for month in months:
                od[month] = format_value_excel(obj.get(month))
            result.append(od)

        for y in (set(years) - set(years_in_db)):
            od = OrderedDict()
            od['year'] = y
            for month in months:
                od[month] = 0.00
            result.append(od)

        result = sorted(result, key=lambda k: int(k['year']))

        wb = xlwt.Workbook(style_compression=2)
        ws = FitSheetWrapper(wb.add_sheet('Monthly YOY report'))

        # Sheet header, first row
        row_num = 0
        bold_center = xlwt.easyxf("font: name Arial, bold True, height 280; align: horiz center")
        ws.row(0).height_mismatch = True
        ws.row(row_num).height = 20 * 20
        ws.write_merge(row_num, 0, 0, 12, airline_name.upper(), bold_center)
        row_num = row_num + 1
        ws.row(row_num).height = 20 * 20
        ws.write_merge(row_num, 1, 0, 12, report_header.upper(), bold_center)
        row_num = row_num + 1
        ws.row(row_num).height = 20 * 20
        ws.write_merge(row_num, 2, 0, 12, ', '.join(years).upper(), bold_center)
        row_num = row_num + 1
        count = 0
        ws.row(row_num).height = 25 * 20
        ws.write(row_num, count, "Year", xlwt.easyxf(
            "align: wrap yes, vert centre, horiz center;font: name Arial, height 180;pattern: pattern solid, fore_colour yellow;border: left thin,right thin,top thin,bottom thin"))
        count = count + 1
        for month in months:
            ws.write(row_num, count, month, xlwt.easyxf(
                "align: wrap yes, vert centre, horiz center;font: name Arial, height 180;pattern: pattern solid, fore_colour yellow;border: left thin,right thin,top thin,bottom thin"))
            count = count + 1
        row_num = row_num + 1
        for row in result:
            count = 0
            for num, ele in row.items():
                ws.row(row_num).height = 20 * 20
                if type(ele) is float:
                    style = xlwt.easyxf(
                        "align: wrap yes, vert centre, horiz right;border: left thin,right thin,top thin,bottom thin")
                    ws.write(row_num, count, format_value_excel(ele), style)
                else:
                    style = xlwt.easyxf(
                        "align: wrap yes, vert centre, horiz left;border: left thin,right thin,top thin,bottom thin")
                    ws.write(row_num, count, ele, style)
                count = count + 1
            row_num = row_num + 1

        wb.save(response)

        return response

    def get_context_data(self, **kwargs):
        context = super(MonthlyYOYReport, self).get_context_data(**kwargs)
        context['activate'] = 'reports'
        context['order_by'] = self.request.GET.get('order_by', 'id')
        context['selected_airline'] = self.request.GET.get('airline', '')
        context['organize_by'] = self.request.GET.get('organize_by', '')
        context['airlines'] = Airline.objects.filter(country=self.request.session.get('country'))
        context['years'] = list(range(2000, datetime.datetime.now().year + 1))
        context['months'] = [calendar.month_abbr[i] for i in range(1, 13)]
        years = self.request.GET.getlist('years')
        years.sort()
        context['selected_years'] = years
        return context


class AirlineAgencyReport(ListView):
    """Sales comparison report."""

    model = Transaction
    template_name = 'airline-agency-report.html'
    context_object_name = 'transactions'
    permission_required = ('report.view_airline_agency',)

    def get_queryset(self):

        organize_by = self.request.GET.get('organize_by', '')
        month_year = self.request.GET.get('month_year', '')
        year_only = self.request.GET.get('year_only', '')

        if is_arc(self.request.session.get('country')):
            qs = Transaction.objects.select_related('agency')
        else:
            qs = Transaction.objects.select_related('agency').exclude(
                Q(transaction_type__startswith='ACM') | Q(transaction_type__startswith='ADM'))

        qs = qs.filter(is_sale=True)

        airlines = self.request.GET.getlist('airlines')
        airline_datas = []

        if airlines:
            qs = qs.filter(report__airline__id__in=airlines)
            airline_datas = Airline.objects.filter(id__in=airlines).values('id', 'name').order_by('name')

        if organize_by == '1':
            month = datetime.datetime.strptime(month_year, '%B %Y').month or ''
            year = datetime.datetime.strptime(month_year, '%B %Y').year or ''
            if month and year:
                # qs = qs.filter(report__report_period__month=month, report__report_period__year=year)
                qs = Transaction.objects.filter(report__report_period__year=year,report__report_period__month=month,report__airline__id__in=airlines).exclude(agency__trade_name="AIRLINEPROS")
        else:
            if year_only:
                qs = qs.filter(report__report_period__year=year_only)

        annotations = {}
        result = {
            "data": [],
            "airline_datas": airline_datas,
            "totals": [],
            "exists": False
        }

        for air in airlines:
            annotation_name = '{}'.format(air)
            if is_arc(self.request.session.get('country')):
                annotations[annotation_name] = Sum('fare_amount', filter=Q(report__airline__id=air))
            else:
                annotations[annotation_name] = Sum('transaction_amount', filter=Q(report__airline__id=air))

        if is_arc(self.request.session.get('country')):
            qs = qs.values('agency__id', 'agency__agency_no', 'agency__sales_owner__email',
                           'agency__trade_name', 'agency__state__abrev', 'agency__tel',
                           'agency__agency_type__name').order_by().distinct().annotate(
                dcount=Count('agency__agency_no'),
                total_sales=Sum(
                    'fare_amount'),
                **annotations)
        else:
            qs = qs.values('agency__agency_no', 'agency__sales_owner__email',
                           'agency__trade_name', 'agency__state__abrev', 'agency__tel',
                           'agency__agency_type__name').order_by().distinct().annotate(
                dcount=Count('agency__agency_no'),
                total_sales=Sum(
                    'transaction_amount'),
                **annotations)

        # hjhj

        od_sum = OrderedDict()
        od_sum["total_sales"] = 0.00
        for elem in airline_datas:
            od_sum[elem.get('id')] = 0.00

        for obj in qs:
            od = OrderedDict()
            od['agency__agency_no'] = obj.get('agency__agency_no')
            od['agency__sales_owner__email'] = obj.get('agency__sales_owner__email')
            od['agency__trade_name'] = obj.get('agency__trade_name')
            od['agency__state__abrev'] = obj.get('agency__state__abrev')
            od['agency__tel'] = obj.get('agency__tel')
            od['agency__agency_type__name'] = obj.get('agency__agency_type__name')
            od['total_sales'] = format_value(obj.get('total_sales'))
            if obj.get('total_sales'):
                od_sum["total_sales"] = od_sum["total_sales"] + obj.get('total_sales')
            for elem in airline_datas:
                od[elem.get('id')] = obj.get(str(elem.get('id')))
                if obj.get(str(elem.get('id'))):
                    od_sum[elem.get('id')] = od_sum[elem.get('id')] + obj.get(str(elem.get('id')))
            result["data"].append(od)
        result["totals"].append(od_sum)
        result["exists"] = qs.exists()
        return result

    def get_context_data(self, **kwargs):
        context = super(AirlineAgencyReport, self).get_context_data(**kwargs)
        context['activate'] = 'reports'
        context['q_string'] = self.request.META['QUERY_STRING']
        context['selected_airline'] = self.request.GET.getlist('airlines')
        context['organize_by'] = self.request.GET.get('organize_by', '')
        context['month_year'] = self.request.GET.get('month_year', '')
        context['year_only'] = self.request.GET.get('year_only', '')
        context['airlines'] = Airline.objects.filter(country=self.request.session.get('country'))
        return context


class GetAirlineAgencyReport(View):
    """Filtered Airline Agency Report download as CSV."""

    permission_required = ('report.download_airline_agency',)

    def get(self, request, *args, **kwargs):
        airline_name = 'AIRLINE NAME'
        response = HttpResponse(content_type='application/vnd.ms-excel')
        report_header = 'AIRLINE AGENCY REPORT'

        month_year_per = "FROM DATE - TO DATE"

        organize_by = self.request.GET.get('organize_by', '')
        month_year = self.request.GET.get('month_year', '')
        year_only = self.request.GET.get('year_only', '')


        if is_arc(self.request.session.get('country')):
            qs = Transaction.objects.select_related('agency')
        else:
            qs = Transaction.objects.select_related('agency').exclude(
                Q(transaction_type__startswith='ACM') | Q(transaction_type__startswith='ADM'))

        qs = qs.filter(is_sale=True)

        airlines = self.request.GET.getlist('airlines')
        airline_datas = []

        if airlines:
            qs = qs.filter(report__airline__id__in=airlines)
            airline_datas = Airline.objects.filter(id__in=airlines).values('id', 'name').order_by('name')
            airline_name = ', '.join(map(lambda x: x['name'], airline_datas))
            response['Content-Disposition'] = 'inline; filename=Airline Agency Report' + '' + '.xls'

        if organize_by == '1':
            month = datetime.datetime.strptime(month_year, '%B %Y').month or ''
            year = datetime.datetime.strptime(month_year, '%B %Y').year or ''
            month_year_per = month_year
            if month and year:
                # qs = qs.filter(report__report_period__month=month, report__report_period__year=year)
                qs = Transaction.objects.filter(report__report_period__year=year,report__report_period__month=month,report__airline__id__in=airlines).exclude(agency__trade_name="AIRLINEPROS")
        else:
            if year_only:
                month_year_per = year_only
                qs = qs.filter(report__report_period__year=year_only)

        annotations = {}
        result = {
            "data": [],
            "airline_datas": airline_datas,
            "totals": []
        }

        for air in airlines:
            annotation_name = '{}'.format(air)
            if is_arc(self.request.session.get('country')):
                annotations[annotation_name] = Sum('fare_amount', filter=Q(report__airline__id=air))
            else:
                annotations[annotation_name] = Sum('transaction_amount', filter=Q(report__airline__id=air))

        

        if is_arc(self.request.session.get('country')):
            qs = qs.values('agency__id', 'agency__agency_no', 'agency__sales_owner__email',
                           'agency__trade_name', 'agency__state__abrev', 'agency__tel',
                           'agency__agency_type__name').order_by().distinct().annotate(
                dcount=Count('agency__agency_no'),
                total_sales=Sum(
                    'fare_amount'),
                **annotations)
        else:
            qs = qs.values('agency__agency_no', 'agency__sales_owner__email',
                           'agency__trade_name', 'agency__state__abrev', 'agency__tel',
                           'agency__agency_type__name').order_by().distinct().annotate(
                dcount=Count('agency__agency_no'),
                total_sales=Sum(
                    'transaction_amount'),
                **annotations)

        od_sum = OrderedDict()
        od_sum["total_sales"] = 0.00
        for elem in airline_datas:
            od_sum[elem.get('id')] = 0.00

        for obj in qs:
            od = OrderedDict()
            od['agency__agency_no'] = obj.get('agency__agency_no')
            od['agency__sales_owner__email'] = obj.get('agency__sales_owner__email')
            od['agency__trade_name'] = obj.get('agency__trade_name')
            od['agency__state__abrev'] = obj.get('agency__state__abrev')
            od['agency__tel'] = obj.get('agency__tel')
            od['agency__agency_type__name'] = obj.get('agency__agency_type__name')
            od['total_sales'] = obj.get('total_sales')
            if obj.get('total_sales'):
                od_sum["total_sales"] = od_sum["total_sales"] + obj.get('total_sales')
            for elem in airline_datas:
                od[elem.get('id')] = obj.get(str(elem.get('id')))
                if obj.get(str(elem.get('id'))):
                    od_sum[elem.get('id')] = od_sum[elem.get('id')] + obj.get(str(elem.get('id')))
            result["data"].append(od)
        result["totals"].append(od_sum)

        wb = xlwt.Workbook(style_compression=2)
        ws = FitSheetWrapper(wb.add_sheet('Airline Agency Report'))
        ws.col(0).width = (15 * 367)

        headers = ["Agency", "Sales Owner", "Agency trade name", "State", "Tel", "Agency type", "Total Sales"]
        for air in airline_datas:
            headers.append(air.get('name'))

        row_num = 0
        bold_center = xlwt.easyxf("font: name Arial, bold True, height 280; align: horiz center")
        ws.row(0).height_mismatch = True
        ws.row(row_num).height = 20 * 20
        ws.write_merge(row_num, 0, 0, len(headers) - 1, airline_name.upper(), bold_center)
        row_num = row_num + 1
        ws.row(row_num).height = 20 * 20
        ws.write_merge(row_num, 1, 0, 7, report_header.upper(), bold_center)
        row_num = row_num + 1
        ws.row(row_num).height = 20 * 20
        ws.write_merge(row_num, 2, 0, 7, month_year_per.upper(), bold_center)
        row_num = row_num + 1

        ws.row(row_num).height = 25 * 20
        count = 0
        for c, h in enumerate(headers):
            ws.write(row_num, count, h, xlwt.easyxf(
                "align: wrap yes, vert centre, horiz center;font: name Arial, height 180, bold True;pattern: pattern solid, fore_colour yellow;border: left thin,right thin,top thin,bottom thin"))
            count = count + 1

        row_num = row_num + 1

        for data in result["data"]:
            count = 0
            for vals in data.values():
                if type(vals) is float:
                    style = xlwt.easyxf(
                        "align: wrap yes, vert centre, horiz right;font: name Arial, height 180;border: left thin,right thin,top thin,bottom thin")
                    ws.write(row_num, count, format_value_excel(vals), style)
                else:
                    style = xlwt.easyxf(
                        "align: wrap yes, vert centre, horiz left;font: name Arial, height 180;border: left thin,right thin,top thin,bottom thin")
                    ws.write(row_num, count, vals, style)

                count = count + 1
            row_num = row_num + 1

        count = 0
        for i in range(0, 5):
            ws.write(row_num, count, " ", xlwt.easyxf(
                "align: wrap yes, vert centre, horiz left;font: name Arial, height 180, bold True;border: left thin,right thin,top thin,bottom thin"))
            count = count + 1

        ws.write(row_num, count, "Total: ", xlwt.easyxf(
            "align: wrap yes, vert centre, horiz left;font: name Arial, height 180, bold True;border: left thin,right thin,top thin,bottom thin"))
        count = count + 1
        for tot in od_sum.values():
            ws.write(row_num, count, tot, xlwt.easyxf(
                "align: wrap yes, vert centre, horiz right;font: name Arial, height 180, bold True;border: left thin,right thin,top thin,bottom thin"))
            count = count + 1

        wb.save(response)
        return response


class SalesComparisonReport(PermissionRequiredMixin, ListView):
    """Sales comparison report."""

    permission_required = ('report.view_sales_comparison',)
    model = Transaction
    template_name = 'sales-comparison-report.html'
    context_object_name = 'transactions'

    def get_queryset(self):
        organize_by = self.request.GET.get('organize_by', '')
        start_month = self.request.GET.get('start_month', '')
        end_month = self.request.GET.get('end_month', '')
        airline = self.request.GET.get('airline', '')
        qs = Transaction.objects.select_related('agency').filter(is_sale=True)
        years = self.request.GET.getlist('years')
        years.sort()

        annotations = {}
        result = []

        for y in years:
            start = datetime.date(year=int(y), month=int(start_month), day=1)
            end = datetime.date(year=int(y), month=int(end_month), day=calendar.monthrange(int(y), int(end_month))[1])
            annotation_name = '{}'.format(y)
            annotations[annotation_name] = Sum('fare_amount', filter=Q(report__report_period__ped__range=[start, end]))

        if airline:
            qs = qs.filter(report__airline=airline)

        if start_month and end_month:
            if organize_by == 'agency':
                qs = qs.values('agency__agency_no', 'agency__sales_owner__email',
                               'agency__trade_name', 'agency__state__abrev', 'agency__tel',
                               'agency__agency_type__name').order_by().distinct().annotate(
                    dcount=Count('agency__agency_no'), **annotations)

                for obj in qs:
                    od = OrderedDict()
                    od['agency__agency_no'] = obj.get('agency__agency_no')
                    od['agency__sales_owner__email'] = obj.get('agency__sales_owner__email')
                    od['agency__trade_name'] = obj.get('agency__trade_name')
                    od['agency__state__abrev'] = obj.get('agency__state__abrev')
                    od['agency__tel'] = obj.get('agency__tel')
                    od['agency__agency_type__name'] = obj.get('agency__agency_type__name')
                    flag_value = False
                    for count, y in enumerate(years):
                        val = format_value(obj.get(y))
                        if float(val.replace(',', '')) != 0.00:
                            flag_value = True

                        if count == 0 or count == 1:
                            od[y] = format_value(obj.get(y))
                        else:
                            if count == 2:
                                od['chng' + str(count)] = self.get_chng(self.get_val(obj.get(years[1])),
                                                                        self.get_val(obj.get(years[0])))
                                od['diff' + str(count)] = format_value(self.get_val(obj.get(years[1])) - self.get_val(
                                    obj.get(years[0])))
                            od[y] = format_value(obj.get(y))
                            od['chng' + str(count + 1)] = self.get_chng(self.get_val(obj.get(years[count])),
                                                                        self.get_val(obj.get(years[count - 1])))
                            od['diff' + str(count + 1)] = format_value(
                                self.get_val(obj.get(years[count])) - self.get_val(
                                    obj.get(years[count - 1])))
                    if len(years) == 2:
                        od['chng' + str(count)] = self.get_chng(self.get_val(obj.get(years[1])),
                                                                self.get_val(obj.get(years[0])))
                        od['diff' + str(count)] = format_value(
                            self.get_val(obj.get(years[1])) - self.get_val(obj.get(years[0])))
                    if flag_value:
                        result.append(od)

            if organize_by == 'state':
                qs = qs.values('agency__state__abrev', 'agency__state__owner__email').order_by().distinct().annotate(
                    dcount=Count('agency__state__abrev'), **annotations)

                for obj in qs:
                    od = OrderedDict()
                    od['agency__state__abrev'] = obj.get('agency__state__abrev')
                    od['agency__state__owner__email'] = obj.get('agency__state__owner__email')
                    flag_value = False
                    for count, y in enumerate(years):
                        val = format_value(obj.get(y))
                        if float(val.replace(',', '')) != 0.00:
                            flag_value = True
                        if count == 0 or count == 1:
                            od[y] = format_value(obj.get(y))
                        else:
                            if count == 2:
                                od['chng' + str(count)] = self.get_chng(self.get_val(obj.get(years[1])),
                                                                        self.get_val(obj.get(years[0])))
                                od['diff' + str(count)] = format_value(self.get_val(obj.get(years[1])) - self.get_val(
                                    obj.get(years[0])))
                            od[y] = format_value(obj.get(y))
                            od['chng' + str(count + 1)] = self.get_chng(self.get_val(obj.get(years[count])),
                                                                        self.get_val(obj.get(years[count - 1])))
                            od['diff' + str(count + 1)] = format_value(
                                self.get_val(obj.get(years[count])) - self.get_val(
                                    obj.get(years[count - 1])))
                    if len(years) == 2:
                        od['chng' + str(count)] = self.get_chng(self.get_val(obj.get(years[1])),
                                                                self.get_val(obj.get(years[0])))
                        od['diff' + str(count)] = format_value(
                            self.get_val(obj.get(years[1])) - self.get_val(obj.get(years[0])))
                    if flag_value:
                        result.append(od)

            if organize_by == 'city':
                qs = qs.values('agency__city__name', 'agency__state__owner__email').order_by().distinct().annotate(
                    dcount=Count('agency__city__name'), **annotations)

                for obj in qs:
                    od = OrderedDict()
                    od['agency__city__name'] = obj.get('agency__city__name')
                    od['agency__state__owner__email'] = obj.get('agency__state__owner__email')
                    flag_value = False
                    for count, y in enumerate(years):
                        val = format_value(obj.get(y))
                        if float(val.replace(',', '')) != 0.00:
                            flag_value = True

                        if count == 0 or count == 1:
                            od[y] = format_value(obj.get(y))
                        else:
                            if count == 2:
                                od['chng' + str(count)] = self.get_chng(self.get_val(obj.get(years[1])),
                                                                        self.get_val(obj.get(years[0])))
                                od['diff' + str(count)] = format_value(self.get_val(obj.get(years[1])) - self.get_val(
                                    obj.get(years[0])))
                            od[y] = format_value(obj.get(y))
                            od['chng' + str(count + 1)] = self.get_chng(self.get_val(obj.get(years[count])),
                                                                        self.get_val(obj.get(years[count - 1])))
                            od['diff' + str(count + 1)] = format_value(
                                self.get_val(obj.get(years[count])) - self.get_val(
                                    obj.get(years[count - 1])))
                    if len(years) == 2:
                        od['chng' + str(count)] = self.get_chng(self.get_val(obj.get(years[1])),
                                                                self.get_val(obj.get(years[0])))
                        od['diff' + str(count)] = format_value(
                            self.get_val(obj.get(years[1])) - self.get_val(obj.get(years[0])))
                    if flag_value:
                        result.append(od)

            if organize_by == 'sales owner':
                qs = qs.values('agency__sales_owner__email').order_by().distinct().annotate(
                    dcount=Count('agency__sales_owner__email'), **annotations)

                for obj in qs:
                    od = OrderedDict()
                    od['agency__sales_owner__email'] = obj.get('agency__sales_owner__email')
                    od['agency__sales_owner__email'] = obj.get('agency__sales_owner__email')
                    flag_value = False
                    for count, y in enumerate(years):
                        val = format_value(obj.get(y))
                        if float(val.replace(',', '')) != 0.00:
                            flag_value = True

                        if count == 0 or count == 1:
                            od[y] = format_value(obj.get(y))
                        else:
                            if count == 2:
                                od['chng' + str(count)] = self.get_chng(self.get_val(obj.get(years[1])),
                                                                        self.get_val(obj.get(years[0])))
                                od['diff' + str(count)] = format_value(self.get_val(obj.get(years[1])) - self.get_val(
                                    obj.get(years[0])))
                            od[y] = format_value(obj.get(y))
                            od['chng' + str(count + 1)] = self.get_chng(self.get_val(obj.get(years[count])),
                                                                        self.get_val(obj.get(years[count - 1])))
                            od['diff' + str(count + 1)] = format_value(
                                self.get_val(obj.get(years[count])) - self.get_val(
                                    obj.get(years[count - 1])))
                    if len(years) == 2:
                        od['chng' + str(count)] = self.get_chng(self.get_val(obj.get(years[1])),
                                                                self.get_val(obj.get(years[0])))
                        od['diff' + str(count)] = format_value(
                            self.get_val(obj.get(years[1])) - self.get_val(obj.get(years[0])))
                    if flag_value:
                        result.append(od)

            if organize_by == 'agency_type':
                qs = qs.values('agency__agency_type__name').order_by().distinct().annotate(
                    dcount=Count('agency__agency_type__name'), **annotations)

                for obj in qs:
                    od = OrderedDict()
                    od['agency__agency_type__name'] = obj.get('agency__agency_type__name')
                    flag_value = False
                    for count, y in enumerate(years):
                        val = format_value(obj.get(y))
                        if float(val.replace(',', '')) != 0.00:
                            flag_value = True

                        if count == 0 or count == 1:
                            od[y] = format_value(obj.get(y))
                        else:
                            if count == 2:
                                od['chng' + str(count)] = self.get_chng(self.get_val(obj.get(years[1])),
                                                                        self.get_val(obj.get(years[0])))
                                od['diff' + str(count)] = format_value(
                                    self.get_val(obj.get(years[1])) - self.get_val(obj.get(years[0])))
                            od[y] = format_value(obj.get(y))
                            od['chng' + str(count + 1)] = self.get_chng(self.get_val(obj.get(years[count])),
                                                                        self.get_val(obj.get(years[count - 1])))
                            od['diff' + str(count + 1)] = format_value(
                                self.get_val(obj.get(years[count])) - self.get_val(
                                    obj.get(years[count - 1])))
                    if len(years) == 2:
                        od['chng' + str(count)] = self.get_chng(self.get_val(obj.get(years[1])),
                                                                self.get_val(obj.get(years[0])))
                        od['diff' + str(count)] = format_value(
                            self.get_val(obj.get(years[1])) - self.get_val(obj.get(years[0])))
                    if flag_value:
                        result.append(od)
        return result

    def get_chng(self, val2, val1):
        result = (val2 - val1) * 100 / val1 if val2 != 0 and val1 != 0 else 0.0
        return format_value(result)


    def get_diff(self, val1, val2):
        return format_value(val2 - val1)

    def get_val(self, val):
        return val if val else 0.00

    def get_context_data(self, **kwargs):
        context = super(SalesComparisonReport, self).get_context_data(**kwargs)
        context['activate'] = 'reports'
        context['q_string'] = self.request.META['QUERY_STRING']
        context['order_by'] = self.request.GET.get('order_by', 'id')
        context['selected_airline'] = self.request.GET.get('airline', '')
        context['selected_state'] = self.request.GET.get('state', '')
        context['organize_by'] = self.request.GET.get('organize_by', '')
        context['start_month'] = self.request.GET.get('start_month', '')
        context['end_month'] = self.request.GET.get('end_month', '')
        context['airlines'] = Airline.objects.filter(country=self.request.session.get('country'))
        context['years'] = list(range(2000, datetime.datetime.now().year + 1))
        years = self.request.GET.getlist('years')
        years.sort()
        context['selected_years'] = years
        return context


class GetSalesComparisonReport(PermissionRequiredMixin, View):
    """Filtered SalesComparison Report download as CSV."""

    permission_required = ('report.download_sales_comparison',)

    def get_chng(self, val2, val1):
        result = (val2 - val1) * 100 / val1 if val2 != 0 and val1 != 0 else 0.0
        return format_value_excel(result)

    def get_diff(self, val1, val2):
        return val2 - val1

    def get_val(self, val):
        return val if val else 0.00

    def gen_excel(self, headers, data, years, ws):
        row_num = 3
        ws.row(row_num).height = 25 * 20
        count = 0
        for c, h in enumerate(headers):
            ws.write(row_num, count, h, xlwt.easyxf(
                "align: wrap yes, vert centre, horiz center;font: name Arial, height 180, bold True;pattern: pattern solid, fore_colour yellow;border: left thin,right thin,top thin,bottom thin"))
            count = count + 1
        for c, y in enumerate(years):
            if c == 0 or c == 1:
                ws.write(row_num, count, y, xlwt.easyxf(
                    "align: wrap yes, vert centre, horiz center;font: name Arial, height 180, bold True;pattern: pattern solid, fore_colour yellow;border: left thin,right thin,top thin,bottom thin"))
                count = count + 1
            else:
                if c == 2:
                    ws.write(row_num, count, "% Change", xlwt.easyxf(
                        "align: wrap yes, vert centre, horiz center;font: name Arial, height 180, bold True;pattern: pattern solid, fore_colour yellow;border: left thin,right thin,top thin,bottom thin"))
                    count = count + 1
                    ws.write(row_num, count, "Difference", xlwt.easyxf(
                        "align: wrap yes, vert centre, horiz center;font: name Arial, height 180, bold True;pattern: pattern solid, fore_colour yellow;border: left thin,right thin,top thin,bottom thin"))

                count = count + 1
                ws.write(row_num, count, y, xlwt.easyxf(
                    "align: wrap yes, vert centre, horiz center;font: name Arial, height 180, bold True;pattern: pattern solid, fore_colour yellow;border: left thin,right thin,top thin,bottom thin"))
                count = count + 1
                ws.write(row_num, count, "% Change", xlwt.easyxf(
                    "align: wrap yes, vert centre, horiz center;font: name Arial, height 180, bold True;pattern: pattern solid, fore_colour yellow;border: left thin,right thin,top thin,bottom thin"))
                count = count + 1
                ws.write(row_num, count, "Difference", xlwt.easyxf(
                    "align: wrap yes, vert centre, horiz center;font: name Arial, height 180, bold True;pattern: pattern solid, fore_colour yellow;border: left thin,right thin,top thin,bottom thin"))

        if len(years) == 2:
            ws.write(row_num, count, "% Change", xlwt.easyxf(
                "align: wrap yes, vert centre, horiz center;font: name Arial, height 180, bold True;pattern: pattern solid, fore_colour yellow;border: left thin,right thin,top thin,bottom thin"))
            count = count + 1
            ws.write(row_num, count, "Difference", xlwt.easyxf(
                "align: wrap yes, vert centre, horiz center;font: name Arial, height 180, bold True;pattern: pattern solid, fore_colour yellow;border: left thin,right thin,top thin,bottom thin"))
        row_num = row_num + 1
        for elem in data:
            count = 0
            for num, ele in elem.items():
                ws.row(row_num).height = 20 * 20
                if type(ele) is float:
                    style = xlwt.easyxf(
                        "align: wrap yes, vert centre, horiz right;border: left thin,right thin,top thin,bottom thin")
                    ws.write(row_num, count, format_value_excel(ele), style)
                else:
                    style = xlwt.easyxf(
                        "align: wrap yes, vert centre, horiz left;border: left thin,right thin,top thin,bottom thin")
                    ws.write(row_num, count, ele, style)
                count = count + 1
            row_num = row_num + 1

        return count

    def get(self, request, *args, **kwargs):
        airline_name = ''
        response = HttpResponse(content_type='application/vnd.ms-excel')

        organize_by = self.request.GET.get('organize_by', '')
        start_month = self.request.GET.get('start_month', '')
        end_month = self.request.GET.get('end_month', '')
        airline = self.request.GET.get('airline', '')
        qs = Transaction.objects.select_related('agency').filter(is_sale=True)
        years = self.request.GET.getlist('years')
        years.sort()
        annotations = {}
        result = []

        for y in years:
            start = datetime.date(year=int(y), month=int(start_month), day=1)
            end = datetime.date(year=int(y), month=int(end_month), day=calendar.monthrange(int(y), int(end_month))[1])
            annotation_name = '{}'.format(y)
            annotations[annotation_name] = Sum('fare_amount', filter=Q(report__report_period__ped__range=[start, end]))

        if airline:
            qs = qs.filter(report__airline=airline)
            airline_data = Airline.objects.get(id=airline)
            airline_name = airline_data.name
            response[
                'Content-Disposition'] = 'inline; filename=' + airline_data.abrev + ' Sales Comparison ' + '' + '.xls'

        if start_month and end_month:
            month = datetime.date(1900, int(start_month), 1).strftime('%B')
            month_end = datetime.date(1900, int(end_month), 1).strftime('%B')
            month_year = month + ' - ' + month_end + ', ' + '/'.join(years)
            if organize_by == 'agency':
                report_header = "SALES COMPARISON REPORT BY AGENCY SALES"
                qs = qs.values('agency__agency_no', 'agency__sales_owner__email',
                               'agency__trade_name', 'agency__state__abrev', 'agency__tel',
                               'agency__agency_type__name').order_by().distinct().annotate(
                    dcount=Count('agency__agency_no'), **annotations)

                for obj in qs:
                    od = OrderedDict()
                    od['agency__agency_no'] = str(obj.get('agency__agency_no'))
                    od['agency__sales_owner__email'] = obj.get('agency__sales_owner__email')
                    od['agency__trade_name'] = obj.get('agency__trade_name')
                    od['agency__state__abrev'] = obj.get('agency__state__abrev')
                    od['agency__tel'] = str(obj.get('agency__tel'))
                    od['agency__agency_type__name'] = obj.get('agency__agency_type__name')
                    flag_value = False
                    for count, y in enumerate(years):
                        val = format_value(obj.get(y))
                        if float(val.replace(',', '')) != 0.00:
                            flag_value = True
                        if count == 0 or count == 1:
                            od[y] = obj.get(y)
                        else:
                            if count == 2:
                                od['chng' + str(count)] = self.get_chng(self.get_val(obj.get(years[1])),
                                                                        self.get_val(obj.get(years[0])))
                                od['diff' + str(count)] = self.get_val(obj.get(years[1])) - self.get_val(
                                    obj.get(years[0]))
                            od[y] = obj.get(y)
                            od['chng' + str(count + 1)] = self.get_chng(self.get_val(obj.get(years[count])),
                                                                        self.get_val(obj.get(years[count - 1])))
                            od['diff' + str(count + 1)] = self.get_val(obj.get(years[count])) - self.get_val(
                                obj.get(years[count - 1]))
                    if len(years) == 2:
                        od['chng' + str(count)] = self.get_chng(self.get_val(obj.get(years[1])),
                                                                self.get_val(obj.get(years[0])))
                        od['diff' + str(count)] = self.get_val(obj.get(years[1])) - self.get_val(obj.get(years[0]))
                    if flag_value:
                        result.append(od)

            if organize_by == 'state':
                report_header = "SALES COMPARISON REPORT BY STATE SALES"
                qs = qs.values('agency__state__abrev', 'agency__state__owner__email').order_by().distinct().annotate(
                    dcount=Count('agency__state__abrev'), **annotations)

                for obj in qs:
                    od = OrderedDict()
                    od['agency__state__abrev'] = obj.get('agency__state__abrev')
                    od['agency__state__owner__email'] = obj.get('agency__state__owner__email')
                    flag_value = False
                    for count, y in enumerate(years):
                        val = format_value(obj.get(y))
                        if float(val.replace(',', '')) != 0.00:
                            flag_value = True
                        if count == 0 or count == 1:
                            od[y] = obj.get(y)
                        else:
                            if count == 2:
                                od['chng' + str(count)] = self.get_chng(self.get_val(obj.get(years[1])),
                                                                        self.get_val(obj.get(years[0])))
                                od['diff' + str(count)] = self.get_val(obj.get(years[1])) - self.get_val(
                                    obj.get(years[0]))
                            od[y] = obj.get(y)
                            od['chng' + str(count + 1)] = self.get_chng(self.get_val(obj.get(years[count])),
                                                                        self.get_val(obj.get(years[count - 1])))
                            od['diff' + str(count + 1)] = self.get_val(obj.get(years[count])) - self.get_val(
                                obj.get(years[count - 1]))
                    if len(years) == 2:
                        od['chng' + str(count)] = self.get_chng(self.get_val(obj.get(years[1])),
                                                                self.get_val(obj.get(years[0])))
                        od['diff' + str(count)] = self.get_val(obj.get(years[1])) - self.get_val(obj.get(years[0]))
                    if flag_value:
                        result.append(od)

            if organize_by == 'city':
                report_header = "SALES COMPARISON REPORT BY CITY SALES"

                qs = qs.values('agency__city__name', 'agency__state__owner__email').order_by().distinct().annotate(
                    dcount=Count('agency__city__name'), **annotations)

                for obj in qs:
                    od = OrderedDict()
                    od['agency__city__name'] = obj.get('agency__city__name')
                    od['agency__state__owner__email'] = obj.get('agency__state__owner__email')
                    flag_value = False
                    for count, y in enumerate(years):
                        val = format_value(obj.get(y))
                        if float(val.replace(',', '')) != 0.00:
                            flag_value = True
                        if count == 0 or count == 1:
                            od[y] = obj.get(y)
                        else:
                            if count == 2:
                                od['chng' + str(count)] = self.get_chng(self.get_val(obj.get(years[1])),
                                                                        self.get_val(obj.get(years[0])))
                                od['diff' + str(count)] = self.get_val(obj.get(years[1])) - self.get_val(
                                    obj.get(years[0]))
                            od[y] = obj.get(y)
                            od['chng' + str(count + 1)] = self.get_chng(self.get_val(obj.get(years[count])),
                                                                        self.get_val(obj.get(years[count - 1])))
                            od['diff' + str(count + 1)] = self.get_val(obj.get(years[count])) - self.get_val(
                                obj.get(years[count - 1]))
                    if len(years) == 2:
                        od['chng' + str(count)] = self.get_chng(self.get_val(obj.get(years[1])),
                                                                self.get_val(obj.get(years[0])))
                        od['diff' + str(count)] = self.get_val(obj.get(years[1])) - self.get_val(obj.get(years[0]))
                    if flag_value:
                        result.append(od)

            if organize_by == 'sales owner':
                report_header = "SALES COMPARISON REPORT BY SALES OWNER SALES"

                qs = qs.values('agency__sales_owner__email').order_by().distinct().annotate(
                    dcount=Count('agency__sales_owner__email'), **annotations)

                for obj in qs:
                    od = OrderedDict()
                    od['agency__sales_owner__email'] = obj.get('agency__sales_owner__email')
                    od['agency__sales_owner__email'] = obj.get('agency__sales_owner__email')
                    flag_value = False
                    for count, y in enumerate(years):
                        val = format_value(obj.get(y))
                        if float(val.replace(',', '')) != 0.00:
                            flag_value = True
                        if count == 0 or count == 1:
                            od[y] = obj.get(y)
                        else:
                            if count == 2:
                                od['chng' + str(count)] = self.get_chng(self.get_val(obj.get(years[1])),
                                                                        self.get_val(obj.get(years[0])))
                                od['diff' + str(count)] = self.get_val(obj.get(years[1])) - self.get_val(
                                    obj.get(years[0]))
                            od[y] = obj.get(y)
                            od['chng' + str(count + 1)] = self.get_chng(self.get_val(obj.get(years[count])),
                                                                        self.get_val(obj.get(years[count - 1])))
                            od['diff' + str(count + 1)] = self.get_val(obj.get(years[count])) - self.get_val(
                                obj.get(years[count - 1]))
                    if len(years) == 2:
                        od['chng' + str(count)] = self.get_chng(self.get_val(obj.get(years[1])),
                                                                self.get_val(obj.get(years[0])))
                        od['diff' + str(count)] = self.get_val(obj.get(years[1])) - self.get_val(obj.get(years[0]))
                    if flag_value:
                        result.append(od)

            if organize_by == 'agency_type':
                report_header = "SALES COMPARISON REPORT BY AGENCY TYPE SALES"

                qs = qs.values('agency__agency_type__name').order_by().distinct().annotate(
                    dcount=Count('agency__agency_type__name'), **annotations)

                for obj in qs:
                    od = OrderedDict()
                    od['agency__agency_type__name'] = obj.get('agency__agency_type__name')
                    flag_value = False
                    for count, y in enumerate(years):
                        val = format_value(obj.get(y))
                        if float(val.replace(',', '')) != 0.00:
                            flag_value = True
                        if count == 0 or count == 1:
                            od[y] = obj.get(y)
                        else:
                            if count == 2:
                                od['chng' + str(count)] = self.get_chng(self.get_val(obj.get(years[1])),
                                                                        self.get_val(obj.get(years[0])))
                                od['diff' + str(count)] = self.get_val(obj.get(years[1])) - self.get_val(
                                    obj.get(years[0]))
                            od[y] = obj.get(y)
                            od['chng' + str(count + 1)] = self.get_chng(self.get_val(obj.get(years[count])),
                                                                        self.get_val(obj.get(years[count - 1])))
                            od['diff' + str(count + 1)] = self.get_val(obj.get(years[count])) - self.get_val(
                                obj.get(years[count - 1]))
                    if len(years) == 2:
                        od['chng' + str(count)] = self.get_chng(self.get_val(obj.get(years[1])),
                                                                self.get_val(obj.get(years[0])))
                        od['diff' + str(count)] = self.get_val(obj.get(years[1])) - self.get_val(obj.get(years[0]))
                    if flag_value:
                        result.append(od)

            wb = xlwt.Workbook(style_compression=2)
            ws = FitSheetWrapper(wb.add_sheet('Sales Comparison'))
            ws.col(0).width = (15 * 367)

            row_num = 3

            if organize_by == 'agency':
                headers = ["Agency", "Sales Owner", "Agency trade name", "State", "Tel", "Agency type"]
                merge_len = self.gen_excel(headers, result, years, ws)

            if organize_by == 'state':
                headers = ["State", "State Owner"]
                merge_len = self.gen_excel(headers, result, years, ws)

            if organize_by == 'city':
                headers = ["City", "State Owner"]
                merge_len = self.gen_excel(headers, result, years, ws)

            if organize_by == 'sales owner':
                headers = ["Sales Owner"]
                merge_len = self.gen_excel(headers, result, years, ws)

            if organize_by == 'agency_type':
                headers = ["Agency Type"]
                merge_len = self.gen_excel(headers, result, years, ws)

            row_num = 0
            bold_center = xlwt.easyxf("font: name Arial, bold True, height 280; align: horiz center")
            ws.row(0).height_mismatch = True
            ws.row(row_num).height = 20 * 20
            ws.write_merge(row_num, 0, 0, merge_len - 1, airline_name.upper(), bold_center)
            row_num = row_num + 1
            ws.row(row_num).height = 20 * 20
            ws.write_merge(row_num, 1, 0, merge_len - 1, report_header.upper(), bold_center)
            row_num = row_num + 1
            ws.row(row_num).height = 20 * 20
            ws.write_merge(row_num, 2, 0, merge_len - 1, month_year.upper(), bold_center)
            row_num = row_num + 1
            wb.save(response)
        return response


class YearToYearSalesReport(PermissionRequiredMixin, View):
    """year to year Sales report."""

    permission_required = ('report.view_year_to_year',)
    model = Transaction
    template_name = 'year-to-year-sales-report.html'
    context_object_name = 'transactions'

    def get(self, request):
        context = dict()
        month_year = self.request.GET.get('month_year', '')
        sales_type = self.request.GET.get('sales_type', '')
        qs = Transaction.objects.select_related('report__report_period', 'report__airline').filter(is_sale=True)
        airlines = Airline.objects.filter(country=self.request.session.get('country'))
        values = []
        is_arc_var = is_arc(self.request.session.get('country'))

        if is_arc_var:
            qs_disp = Disbursement.objects.select_related('report_period', 'airline')
        else:
            qs_acm = Transaction.objects.select_related('report__report_period', 'report__airline', 'agency').filter(
                transaction_type__startswith='ACM', agency__agency_no='6998051', is_sale=True)

        if month_year:

            month = datetime.datetime.strptime(month_year, '%B %Y').month or ''
            year = datetime.datetime.strptime(month_year, '%B %Y').year or ''
            context['month_year2'] = datetime.datetime(year, month, 1).strftime("%b %Y")
            context['month_year1'] = (datetime.datetime(year, month, 1) - rd(years=1)).strftime("%b %Y")
            if month and year:
                qs1 = qs.filter(report__report_period__month=month, report__report_period__year=year - 1)
                qs2 = qs.filter(report__report_period__month=month, report__report_period__year=year)
                if is_arc_var:
                    qs1_disp = qs_disp.filter(report_period__month=month, report_period__year=year - 1)
                    qs2_disp = qs_disp.filter(report_period__month=month, report_period__year=year)
                else:
                    qs1_acm = qs_acm.filter(report__report_period__month=month, report__report_period__year=year - 1)
                    qs2_acm = qs_acm.filter(report__report_period__month=month, report__report_period__year=year)

                airlines = airlines.filter(
                    id__in=qs.filter(report__report_period__year__in=[year - 1, year],
                                     report__report_period__month=month).values_list(
                        'report__airline', flat=True))
                for airline in airlines:
                    sales_data = {}
                    sales_data['airline'] = airline

                    if sales_type == 'Net':
                        if is_arc_var:
                            acm_y1 = Transaction.objects.select_related('report', 'report__airline',
                                                                        'report__report_period').filter(
                                report__airline=airline, transaction_type__startswith='ACM',
                                agency__agency_no='31768026', report__report_period__month=month,
                                report__report_period__year=year - 1).aggregate(
                                total_ap_acm=Coalesce(Sum('transaction_amount'), V(0))).get('total_ap_acm')
                            acm_y2 = Transaction.objects.select_related('report', 'report__airline',
                                                                        'report__report_period').filter(
                                report__airline=airline, transaction_type__startswith='ACM',
                                agency__agency_no='31768026', report__report_period__month=month,
                                report__report_period__year=year).aggregate(
                                total_ap_acm=Coalesce(Sum('transaction_amount'), V(0))).get(
                                'total_ap_acm')

                            sales_data['amount1'] = qs1.filter(report__airline=airline).aggregate(
                                amount1=Coalesce(Sum('fare_amount'), V(0)) + Coalesce(Sum('pen'), V(0))).get(
                                'amount1') - acm_y1 or 0
                            sales_data['cash1'] = qs1_disp.filter(airline=airline).aggregate(
                                cash1=Sum('bank7')).get('cash1') or 0
                            sales_data['amount2'] = qs2.filter(report__airline=airline).aggregate(
                                amount2=Coalesce(Sum('fare_amount'), V(0)) + Coalesce(Sum('pen'), V(0))).get(
                                'amount2') - acm_y2 or 0
                            sales_data['cash2'] = qs2_disp.filter(airline=airline).aggregate(
                                cash1=Sum('bank7')).get('cash1') or 0
                        else:
                            qs1_acm_nw = qs1_acm.filter(report__airline=airline).aggregate(
                                total=Sum('fare_amount', output_field=FloatField()),
                                total_comm=Sum('std_comm_amount', output_field=FloatField()),
                                total_cp=Sum('pen', output_field=FloatField()))

                            qs2_acm_nw = qs2_acm.filter(report__airline=airline).aggregate(
                                total=Sum('fare_amount', output_field=FloatField()),
                                total_comm=Sum('std_comm_amount', output_field=FloatField()),
                                total_cp=Sum('pen', output_field=FloatField()))
                            sales_data['amount1'] = qs1.filter(report__airline=airline).aggregate(
                                amount1=(Sum('fare_amount') - Coalesce(qs1_acm_nw.get('total'), V(0))) - (
                                    Sum('std_comm_amount', output_field=FloatField()) - Coalesce(
                                    qs1_acm_nw.get('total_comm'), V(0))) + (
                                            Sum('pen') - Coalesce(qs1_acm_nw.get('total_cp'), V(0)))).get(
                                'amount1') or 0
                            sales_data['cash1'] = qs1.filter(report__airline=airline).aggregate(
                                cash1=Sum('ca') - (Sum('std_comm_amount', output_field=FloatField()) - Coalesce(
                                    qs1_acm_nw.get('total_comm'), V(0)))).get(
                                'cash1') or 0
                            sales_data['amount2'] = qs2.filter(report__airline=airline).aggregate(
                                amount2=(Sum('fare_amount') - Coalesce(qs2_acm_nw.get('total'), V(0))) - (
                                    Sum('std_comm_amount', output_field=FloatField()) - Coalesce(
                                    qs2_acm_nw.get('total_comm'), V(0))) + (
                                            Sum('pen') - Coalesce(qs2_acm_nw.get('total_cp'), V(0)))).get(
                                'amount2') or 0
                            sales_data['cash2'] = qs2.filter(report__airline=airline).aggregate(
                                cash2=Sum('ca') - (Sum('std_comm_amount', output_field=FloatField()) - Coalesce(
                                    qs2_acm_nw.get('total_comm'), V(0)))).get(
                                'cash2') or 0
                    else:
                        if is_arc_var:
                            sales_data['amount1'] = qs1.filter(report__airline=airline).aggregate(
                                amount1=Sum('transaction_amount', output_field=FloatField()) - Sum('std_comm_amount',
                                                                                                   output_field=FloatField())).get(
                                'amount1') or 0
                            sales_data['amount2'] = qs2.filter(report__airline=airline).aggregate(
                                amount2=Sum('transaction_amount', output_field=FloatField()) - Sum('std_comm_amount',
                                                                                                   output_field=FloatField())).get(
                                'amount2') or 0
                            sales_data['cash1'] = qs1_disp.filter(airline=airline).aggregate(
                                cash1=Sum('bank7')).get('cash1') or 0
                            sales_data['cash2'] = qs2_disp.filter(airline=airline).aggregate(
                                cash1=Sum('bank7')).get('cash1') or 0
                        else:
                            qs1_acm_nw = qs1_acm.filter(report__airline=airline).aggregate(
                                total=Sum('transaction_amount', output_field=FloatField()),
                                total_comm=Sum('std_comm_amount', output_field=FloatField()))

                            qs2_acm_nw = qs2_acm.filter(report__airline=airline).aggregate(
                                total=Sum('transaction_amount', output_field=FloatField()),
                                total_comm=Sum('std_comm_amount', output_field=FloatField()))

                            sales_data['amount1'] = qs1.filter(report__airline=airline).aggregate(
                                amount1=(Sum('transaction_amount', output_field=FloatField()) - Coalesce(
                                    qs1_acm_nw.get('total'), V(0))) - (
                                            Sum('std_comm_amount', output_field=FloatField()) - Coalesce(
                                            qs1_acm_nw.get('total_comm'), V(0)))).get(
                                'amount1') or 0
                            sales_data['amount2'] = qs2.filter(report__airline=airline).aggregate(
                                amount2=(Sum('transaction_amount', output_field=FloatField()) - Coalesce(
                                    qs2_acm_nw.get('total'), V(0))) - (
                                            Sum('std_comm_amount', output_field=FloatField()) - Coalesce(
                                            qs2_acm_nw.get('total_comm'), V(0)))).get(
                                'amount2') or 0
                            sales_data['cash1'] = qs1.filter(report__airline=airline).aggregate(
                                cash1=Sum('ca') - (Sum('std_comm_amount', output_field=FloatField()) - Coalesce(
                                    qs1_acm_nw.get('total_comm'), V(0)))).get(
                                'cash1') or 0
                            sales_data['cash2'] = qs2.filter(report__airline=airline).aggregate(
                                cash2=Sum('ca') - (Sum('std_comm_amount', output_field=FloatField()) - Coalesce(
                                    qs2_acm_nw.get('total_comm'), V(0)))).get(
                                'cash2') or 0

                    values.append(sales_data)
                context['values'] = values

        context['activate'] = 'reports'
        context['sales_type'] = self.request.GET.get('sales_type', '')
        context['month_year'] = self.request.GET.get('month_year', '')
        context['values'] = values
        context['is_arc'] = is_arc(self.request.session.get('country'))
        return render(request, self.template_name, context)


class CommissionReport(PermissionRequiredMixin, ListView):
    """Commission report."""

    permission_required = ('report.view_commission',)
    model = Transaction
    template_name = 'commission-report.html'
    context_object_name = 'transactions'
    paginate_by = 1000

    def get_queryset(self):
        month_year = self.request.GET.get('month_year', '')
        start_date = self.request.GET.get('start_date', '')
        end_date = self.request.GET.get('end_date', '')
        airline = self.request.GET.get('airline', '')
        trns_list = []

        if airline:
            qs = Transaction.objects.select_related('agency__city', 'agency__state', 'report__airline').exclude(
                std_comm_rate__isnull=True).order_by('report__report_period__ped', 'report__airline').filter(
                report__airline=airline)

            if month_year:

                month = datetime.datetime.strptime(month_year, '%B %Y').month or ''
                year = datetime.datetime.strptime(month_year, '%B %Y').year or ''
                if month and year:
                    qs = qs.filter(report__report_period__month=month, report__report_period__year=year)

            if start_date and end_date:
                start = datetime.datetime.strptime(start_date, '%d %B %Y')
                end = datetime.datetime.strptime(end_date, '%d %B %Y')
                qs = qs.filter(report__report_period__ped__range=[start, end])

            trns_list = []
            commissions_history = CommissionHistory.objects.filter(airline=airline, type='M').values('from_date',
                                                                                                     'to_date',
                                                                                                     'rate').order_by(
                '-from_date')
            for transaction in qs:
                histories = [history for history in commissions_history if
                             history.get('from_date') <= transaction.report.report_period.ped]
                rate = histories[0].get('rate') or 0.0 if histories else 0.0
                if transaction.std_comm_rate < rate:
                    trn = dict()
                    trn['agency_no'] = transaction.agency.agency_no
                    trn['agency_id'] = transaction.agency.pk
                    trn['agency_name'] = transaction.agency.trade_name
                    trn['agency_address'] = transaction.agency.address1
                    trn['agency_city'] = transaction.agency.city.name if transaction.agency.city else ''
                    trn['agency_state'] = transaction.agency.state.name if transaction.agency.state else ''
                    trn['agency_tel'] = transaction.agency.tel
                    trn['ticket_no'] = transaction.ticket_no
                    trn['ped'] = transaction.report.report_period.ped
                    trn['std_comm_rate'] = transaction.std_comm_rate
                    trn['max_comm_rate'] = rate
                    trns_list.append(trn)

        return trns_list

    def get_context_data(self, **kwargs):
        context = super(CommissionReport, self).get_context_data(**kwargs)
        context['activate'] = 'reports'
        context['month_year'] = self.request.GET.get('month_year', '')
        context['selected_airline'] = self.request.GET.get('airline', '')
        context['date_filter'] = self.request.GET.get('date_filter', 'month_year')
        context['start_date'] = self.request.GET.get('start_date', '')
        context['end_date'] = self.request.GET.get('end_date', '')
        context['airlines'] = Airline.objects.filter(country=self.request.session.get('country'))
        return context


class TopAgentReport(PermissionRequiredMixin, ListView):
    """top agent report."""

    permission_required = ('report.view_top_agency',)
    model = Transaction
    template_name = 'top-agent-report.html'
    context_object_name = 'transactions'
    paginate_by = 100

    def get_queryset(self):
        start_date = self.request.GET.get('start_date', '')
        end_date = self.request.GET.get('end_date', '')
        airline = self.request.GET.get('airline', '')
        selected_city = self.request.GET.get('city', '')
        selected_state = self.request.GET.get('state', '')
        limit = int(self.request.GET.get('limit', 1))
        qs = Transaction.objects.select_related('agency__city', 'agency__state', 'report__airline').exclude(
            transaction_type__startswith='ACM', agency__agency_no='6998051')

        qs = qs.filter(is_sale=True)

        if airline:
            qs = qs.filter(report__airline=airline)

        if start_date and end_date:
            start = datetime.datetime.strptime(start_date, '%d %B %Y')
            end = datetime.datetime.strptime(end_date, '%d %B %Y')
            qs = qs.filter(report__report_period__ped__range=[start, end])

        if selected_state:
            qs = qs.filter(agency__state=selected_state)

        if selected_city:
            qs = qs.filter(agency__city=selected_city)

        return qs.values('agency').annotate(
            fare_sum=Coalesce(Sum('fare_amount'), V(0)) - Coalesce(Sum('std_comm_amount'), V(0)) + Coalesce(Sum('pen'),
                                                                                                            V(0)),
            agency_name=F('agency__trade_name'),
            agency_city=F('agency__city__name'), agency_state=F('agency__state__name'),
            tel=F('agency__tel'), agency_email=F('agency__email'),
            agency_no=F('agency__agency_no'),
            sales_owner=F('agency__sales_owner__email')).order_by('-fare_sum')[:limit]

    def get_context_data(self, **kwargs):
        context = super(TopAgentReport, self).get_context_data(**kwargs)
        context['activate'] = 'reports'
        context['limit'] = self.request.GET.get('limit', 1)
        context['selected_airline'] = self.request.GET.get('airline', '')
        context['start_date'] = self.request.GET.get('start_date', '')
        context['end_date'] = self.request.GET.get('end_date', '')
        context['airlines'] = Airline.objects.filter(country=self.request.session.get('country')).order_by('name')
        context['states'] = State.objects.filter(country=self.request.session.get('country')).order_by('name')
        context['cities'] = City.objects.filter(country=self.request.session.get('country')).order_by('name')
        context['selected_city'] = self.request.GET.get('city', '')
        context['selected_state'] = self.request.GET.get('state', '')
        context['filter'] = self.request.GET.get('filter', 'state')
        return context



def adm_report_ignorelist(request,**kwargs):
    import json
    permission_required = ('report.view_adm',)


    if request.is_ajax() and request.method == 'GET' :

        id = request.GET.get('agency_no',' ')

        airline = request.GET.get('airline', '')
        month_year = request.GET.get('monthyear','')
        month_year = month_year.rstrip()
        date= None

        if month_year:
            date= datetime.datetime.strptime( month_year,'%B %Y')


        ex_list=[]
        if id and id !=0:
            ex=Agency.objects.filter(agency_collection_id=id)
            for x in ex:
               ex_list.append(x.agency_no)


        responsee = json.dumps(ex_list)

        return HttpResponse(responsee,content_type='application/json')
    return HttpResponse(None)




class ADMReport(PermissionRequiredMixin, ListView):
    """ADM listing with pagination."""

    permission_required = ('report.view_adm',)
    model = AgencyDebitMemo
    template_name = 'adm-reports.html'
    context_object_name = 'adms'
    missing_dates_ped = ()
    missing_dates_details = ()
    missing_dates_disbursement = ()
    is_arc_var = False

    def get_queryset(self):
        month_year = self.request.GET.get('month_year', '')
        airline = self.request.GET.get('airline', '')
        country = self.request.session.get('country')
        self.is_arc_var = is_arc(self.request.session.get('country'))

        adms = []
        if airline or month_year:

            if not self.is_arc_var:
                trans_type = 'TKTT'
            else:
                trans_type = 'TKT'

            qs = Transaction.objects.filter(transaction_type=trans_type)
            adm_objs = AgencyDebitMemo.objects.select_related('transaction').filter(
                transaction__transaction_type=trans_type)
            if airline:
                qs = qs.filter(report__airline=airline)
                #

                adm_objs = adm_objs.filter(transaction__report__airline=airline)
            if month_year:
                month = datetime.datetime.strptime(month_year, '%B %Y').month or ''
                year = datetime.datetime.strptime(month_year, '%B %Y').year or ''
                month_date = datetime.datetime.strptime(month_year, '%B %Y')
                if month and year:
                    qs = qs .filter(report__report_period__month=month,
                                   report__report_period__year=year)
                    #

                    adm_objs = adm_objs.filter(transaction__report__report_period__month=month,
                                           transaction__report__report_period__year=year)
                    #

                    if qs.count() + adm_objs.count() > 10000:
                        # 10000:
                        base_url = self.request.scheme + '://' + self.request.get_host()
                        excel_adm_report(country, month_year, airline, self.request.user.email, base_url)
                        messages.add_message(self.request, messages.WARNING,
                                             'Excel file generation is taking more time than expected. You will receive an email with the link to download the file once it is done.')
                        return HttpResponseRedirect('/reports/adm/?' + self.request.META.get('QUERY_STRING'))

                    adms_list = adm_objs.values_list('transaction__ticket_no', flat=True)
                    # ---------------------------------------------------------------------------


                    allowed_commission_rate = 0.00

                    for obj in qs:
                        if obj.ticket_no not in adms_list  :
                            #  and obj.agency.agency_no not in exclude_list
                            allowed_commission_rate = 0.00

                            commission_history = CommissionHistory.objects.filter(airline=airline, type='M',
                                                                                  from_date__lte=obj.report.report_period.ped)
                            if commission_history.exists():
                                allowed_commission_rate = commission_history.order_by('-from_date').first().rate

                            try:
                                taken_commission_rate = obj.std_comm_rate
                                commission_rate_diff = allowed_commission_rate - float(taken_commission_rate)
                                cobl_amount = obj.fare_amount
                                if cobl_amount:
                                    if commission_rate_diff < 0:
                                        # ADM
                                        adm_amount = (abs(commission_rate_diff) * cobl_amount) / 100
                                        comment = "Commission deducted " + str(
                                            taken_commission_rate) + "%. Carrier authorized " + str(
                                            allowed_commission_rate) + "%"
                                        allowed_commission_amount = (cobl_amount * allowed_commission_rate) / 100
                                        adms.append({
                                            'agency_no': obj.agency.agency_no,
                                            'trade_name': obj.agency.trade_name,
                                            'ticket_no': obj.ticket_no,
                                            'issue_date': obj.issue_date,
                                            'fare_amount': obj.fare_amount,
                                            'std_comm_amount': obj.std_comm_amount,
                                            'std_comm_rate': obj.std_comm_rate,
                                            'allowed_commission_amount': allowed_commission_amount,
                                            'amount': adm_amount,
                                            'comment': comment

                                        })

                            except Exception as e:
                                pass

                    for obj in adm_objs:
                        # if obj.transaction.agency.agency.agency_no not in exclude_list:
                            adms.append({
                                'agency_no': obj.transaction.agency.agency_no,
                                'trade_name': obj.transaction.agency.trade_name,
                                'ticket_no': obj.transaction.ticket_no,
                                'issue_date': obj.transaction.issue_date,
                                'fare_amount': obj.transaction.fare_amount,
                                'std_comm_amount': obj.transaction.std_comm_amount,
                                'std_comm_rate': obj.transaction.std_comm_rate,
                                'allowed_commission_amount': obj.allowed_commission_amount,
                                'amount': obj.amount,
                                'comment': obj.comment
                            })

                    months_weekend = ReportPeriod.objects.filter(year=year, month=month, country=country).values_list(
                        'ped', flat=True).order_by('ped')

                    if months_weekend:
                        details_peds = ReportFile.objects.filter(report_period__ped__year=year,
                                                                 report_period__ped__month=month,
                                                                 airline=airline, country=country).values_list(
                            'report_period__ped', flat=True)

                        first_ped = months_weekend[0]
                        last_ped = months_weekend[len(months_weekend) - 1]
                        start = first_ped - datetime.timedelta(days=5)
                        end = last_ped + datetime.timedelta(days=1)
                        days = [start + datetime.timedelta(day) for day in range((end - start).days + 1)]
                        if not self.is_arc_var:
                            details_daily = DailyCreditCardFile.objects.filter(airline=airline, date__gte=start,
                                                                               date__lte=end).values_list(
                                'date', flat=True)
                            self.disbursement_files = []
                        else:
                            details_daily = CarrierDeductions.objects.filter(airline=airline, filedate__gte=start,
                                                                             filedate__lte=end).values_list(
                                'filedate', flat=True)

                            self.disbursement_files = Disbursement.objects.filter(airline=airline, filedate__gte=start,
                                                                                  filedate__lte=end).values_list(
                                'filedate', flat=True)

                        self.missing_dates_ped = set(months_weekend) - set(details_peds)

                        if self.is_arc_var:
                            self.missing_dates_disbursement = set(months_weekend) - set(self.disbursement_files)
                            missing_dates_credit = ()
                            if months_weekend.exists():
                                missing_dates_credit_val = months_weekend[0]
                                if missing_dates_credit_val not in details_daily:
                                    missing_dates_credit = [missing_dates_credit_val]
                            self.missing_dates_details = missing_dates_credit
                        else:
                            self.missing_dates_disbursement = ()
                            self.missing_dates_details = set(days) - set(details_daily)
        return adms

    def get_context_data(self, **kwargs):
        context = super(ADMReport, self).get_context_data(**kwargs)
        context['activate'] = 'reports'
        airline = self.request.GET.get('airline', '')
        allowed_commission_rate = 0.00

        month_year = self.request.GET.get('month_year', '')
        if month_year:
            month_date = datetime.datetime.strptime(month_year, '%B %Y')

            commission_history = CommissionHistory.objects.filter(airline=airline, type='M', from_date__lte=month_date)
            if commission_history.exists():
                allowed_commission_rate = commission_history.order_by('-from_date').first().rate


        context['allowed_commission_rate'] = allowed_commission_rate
        context['exclude_list'] =AgencyCollection.objects.all()
        #context['exclude_list'] =AgencyCollection.objects.filter(country__name= self.request.session.get('country'))
        context['month_year'] = month_year
        context['selected_airline'] = self.request.GET.get('airline', '')
        context['missing_dates_ped'] = sorted(self.missing_dates_ped)
        context['missing_dates_details'] = sorted(self.missing_dates_details)
        context['missing_dates_disbursement'] = sorted(self.missing_dates_disbursement)
        context['missing_dates_count'] = len(self.missing_dates_ped) + len(self.missing_dates_details) + len(
            self.missing_dates_disbursement)
        context['airlines'] = Airline.objects.filter(country=self.request.session.get('country'))
        return context


def excel_adm_report(country, month_year, airline, to_email, base_url,*args,**kwargs):
    import threading
    agency_list=kwargs.get('agency_list','')
    exclude_list=[]

    ex=Agency.objects.filter(agency_collection_id=agency_list)
    if ex:
        for x in ex:
            exclude_list.append(x.agency_no)

    def thread_process():
        airline_name = ''
        allowed_commission_rate = 0.00

        adms = []
        if airline or month_year:
            if Country.objects.get(id=country).name != 'United States':
                trans_type = 'TKTT'
            else:
                trans_type = 'TKT'

            qs = Transaction.objects.filter(transaction_type=trans_type)
            adm_objs = AgencyDebitMemo.objects.select_related('transaction').filter(
                transaction__transaction_type=trans_type)
            if airline:
                qs = qs.filter(report__airline=airline)
                adm_objs = adm_objs.filter(transaction__report__airline=airline)
                airline_data = Airline.objects.get(id=airline)
                airline_name = airline_data.name
            if month_year:
                month = datetime.datetime.strptime(month_year, '%B %Y').month or ''
                year = datetime.datetime.strptime(month_year, '%B %Y').year or ''
                month_date = datetime.datetime.strptime(month_year, '%B %Y')
                if month and year:

                    qs = qs.filter(report__report_period__month=month,
                                   report__report_period__year=year)

                    adm_objs = adm_objs.filter(transaction__report__report_period__month=month,
                                               transaction__report__report_period__year=year)

                    adms_list = adm_objs.values_list('transaction__ticket_no', flat=True)

                    allowed_commission_rate = 0.00
                    file_name = airline_data.abrev + ' ' + month_year + ' ADM Report' + '.xls'

                    for obj in qs:
                        if obj.ticket_no not in adms_list and obj.agency.agency_no not in exclude_list:
                            allowed_commission_rate = 0.00

                            commission_history = CommissionHistory.objects.filter(airline=airline, type='M',
                                                                                  from_date__lte=obj.report.report_period.ped)
                            if commission_history.exists():
                                allowed_commission_rate = commission_history.order_by('-from_date').first().rate

                            try:
                                taken_commission_rate = obj.std_comm_rate
                                commission_rate_diff = allowed_commission_rate - float(taken_commission_rate)
                                cobl_amount = obj.fare_amount
                                if cobl_amount:
                                    if commission_rate_diff < 0:
                                        # ADM
                                        adm_amount = (abs(commission_rate_diff) * cobl_amount) / 100
                                        comment = "Commission deducted " + str(
                                            taken_commission_rate) + "%. Carrier authorized " + str(
                                            allowed_commission_rate) + "%"
                                        allowed_commission_amount = (cobl_amount * allowed_commission_rate) / 100
                                        adms.append({
                                            'agency_no': obj.agency.agency_no,
                                            'trade_name': obj.agency.trade_name,
                                            'ticket_no': obj.ticket_no,
                                            'issue_date': obj.issue_date,
                                            'fare_amount': obj.fare_amount,
                                            'std_comm_amount': obj.std_comm_amount,
                                            'std_comm_rate': obj.std_comm_rate,
                                            'allowed_commission_amount': allowed_commission_amount,
                                            'amount': adm_amount,
                                            'comment': comment

                                        })

                            except Exception as e:
                                pass

                    for obj in adm_objs:
                        if obj.transaction.agency.agency.agency_no not in exclude_list:
                            adms.append({
                                'agency_no': obj.transaction.agency.agency_no,
                                'trade_name': obj.transaction.agency.trade_name,
                                'ticket_no': obj.transaction.ticket_no,
                                'issue_date': obj.transaction.issue_date,
                                'fare_amount': obj.transaction.fare_amount,
                                'std_comm_amount': obj.transaction.std_comm_amount,
                                'std_comm_rate': obj.transaction.std_comm_rate,
                                'allowed_commission_amount': obj.allowed_commission_amount,
                                'amount': obj.amount,
                                'comment': obj.comment
                            })

        wb = xlwt.Workbook(style_compression=2)
        ws = wb.add_sheet('ADMS')

        # Sheet header, first row
        row_num = 0

        bold_center = xlwt.easyxf("align: wrap yes, vert centre, horiz center;font: name Arial,height 280, bold True")
        wrap_data = xlwt.easyxf("align: wrap yes, vert centre, horiz center;font: name Arial,height 180")

        ws.row(row_num).height = 20 * 20
        ws.write_merge(row_num, 0, 0, 10, airline_name.upper(), bold_center)
        row_num = row_num + 1
        ws.row(row_num).height = 20 * 20
        ws.write_merge(row_num, 1, 0, 10, "ADM Reports".upper(), bold_center)
        row_num = row_num + 1
        ws.row(row_num).height = 20 * 20
        ws.write_merge(row_num, 2, 0, 10, month_year.upper(), bold_center)
        row_num = row_num + 1
        ws.row(row_num).height_mismatch = True
        ws.row(row_num).height = 20 * 36

        ws.col(0).width = (256 * 13)
        ws.col(1).width = (256 * 35)
        ws.col(2).width = (256 * 12)
        ws.col(3).width = (256 * 12)
        ws.col(4).width = (256 * 11)
        ws.col(5).width = (256 * 12)
        ws.col(6).width = (256 * 13)
        ws.col(7).width = (256 * 14)
        ws.col(8).width = (256 * 11)
        ws.col(9).width = (256 * 40)
        ws.col(10).width = (256 * 9)

        columns = ["Agency Number", "Agency Name", "Tkt Number", "Issue Date", "Tkt Base Fare", "Deducted Commission",
                   "Commission Charged( %)", "Authorized Commission " + str(allowed_commission_rate) + " %",
                   "ADM Amount", "Comments", "ADM NO"]

        for col_num in range(len(columns)):
            ws.write(row_num, col_num, columns[col_num], xlwt.easyxf(
                "align: wrap yes, vert centre, horiz center;font: name Arial, bold True, height 180;pattern: pattern solid,fore-colour yellow;border: left thin,right thin,top thin,bottom thin"))


        for d in adms:
            row_num = row_num + 1
            ws.row(row_num).height = 256 * 3
            col_num = 0
            ws.write(row_num, col_num, d.get('agency_no'), wrap_data)
            col_num = col_num + 1
            ws.write(row_num, col_num, d.get('trade_name'), wrap_data)
            col_num = col_num + 1
            ws.write(row_num, col_num, d.get('ticket_no'), wrap_data)
            col_num = col_num + 1
            ws.write(row_num, col_num, d.get('issue_date').strftime("%Y-%m-%d"), wrap_data)
            col_num = col_num + 1
            ws.write(row_num, col_num, d.get('fare_amount'), wrap_data)
            col_num = col_num + 1
            ws.write(row_num, col_num, d.get('std_comm_amount'), wrap_data)
            col_num = col_num + 1
            ws.write(row_num, col_num, d.get('std_comm_rate'), wrap_data)
            col_num = col_num + 1
            ws.write(row_num, col_num, d.get('allowed_commission_amount'), wrap_data)
            col_num = col_num + 1
            ws.write(row_num, col_num, d.get('amount'), wrap_data)
            col_num = col_num + 1
            ws.write(row_num, col_num, d.get('comment'), wrap_data)
            col_num = col_num + 1
            ws.write(row_num, col_num, '', wrap_data)
            col_num = col_num + 1

        path = os.path.join(settings.MEDIA_ROOT, 'excelreports', file_name)
        wb.save(path)

        excel = ExcelReportDownload()
        excel.report_type = 3
        excel.file.name = 'excelreports/' + file_name
        excel.save()
        context = {'user': to_email,
                   'link': base_url + excel.file.url,
                   'report_name': 'ADM Report',
                   'airline_name': airline_name,
                   'date': month_year
                   }
        send_mail('ADM Report ' + month_year + ': ' + airline_name, "email/excel-report-email.html", context,
                  [to_email], from_email='assda@assda.com')

    t = threading.Thread(target=thread_process())
    t.start()
    return True


class GetADMReport(PermissionRequiredMixin, View):
    """Filtered ADM Report download as CSV."""

    permission_required = ('report.download_adm',)

    def get(self, request, *args, **kwargs):
        month_year = self.request.GET.get('month_year', '')
        airline = self.request.GET.get('airline', '')
        country = self.request.session.get('country')
        agency_list=self.request.GET.get('agency_list','')

        airline_name = ''
        allowed_commission_rate = 0.00

        response = HttpResponse(content_type='application/vnd.ms-excel')
        adms = []
        exclude_list=[]

        ex=Agency.objects.filter(agency_collection_id=agency_list)
        if ex:
            for x in ex:
                exclude_list.append(x.agency_no)


        if airline or month_year:
            if Country.objects.get(id=country).name == 'Canada':
                trans_type = 'TKTT'
            else:
                trans_type = 'TKT'

            qs = Transaction.objects.filter(transaction_type=trans_type)
            adm_objs = AgencyDebitMemo.objects.select_related('transaction').filter(
                transaction__transaction_type=trans_type)


            if qs.count() + adm_objs.count() > 10000:
                #
                base_url = request.scheme + '://' + request.get_host()

                excel_adm_report(country, month_year, airline, request.user.email, base_url,agency_list=agency_list)
                messages.add_message(self.request, messages.WARNING,
                                     'Excel file generation is taking more time than expected. You will receive an email with the link to download the file once it is done.')
                return HttpResponseRedirect('/reports/adm/?' + request.META.get('QUERY_STRING'))
            else:
                if airline:
                    qs = qs.filter(report__airline=airline)
                    adm_objs = adm_objs.filter(transaction__report__airline=airline)
                    airline_data = Airline.objects.get(id=airline)
                    airline_name = airline_data.name
                if month_year:
                    month = datetime.datetime.strptime(month_year, '%B %Y').month or ''
                    year = datetime.datetime.strptime(month_year, '%B %Y').year or ''
                    month_date = datetime.datetime.strptime(month_year, '%B %Y')
                    if month and year:

                        qs = qs .filter(report__report_period__month=month,
                                       report__report_period__year=year)
                        #

                        adm_objs = adm_objs.filter(transaction__report__report_period__month=month,
                                         transaction__report__report_period__year=year)
                        #
                        adms_list = adm_objs.values_list('transaction__ticket_no', flat=True)

                        allowed_commission_rate = 0.00

                        response[
                            'Content-Disposition'] = 'inline; filename=' + airline_data.abrev + ' ' + month_year + ' ADM Report' + '.xls'

                        for obj in qs:
                            if obj.ticket_no not in adms_list and obj.agency.agency_no not in exclude_list:
                                allowed_commission_rate = 0.00

                                commission_history = CommissionHistory.objects.filter(airline=airline, type='M',
                                                                                      from_date__lte=obj.report.report_period.ped)
                                if commission_history.exists():
                                    allowed_commission_rate = commission_history.order_by('-from_date').first().rate

                                try:
                                    taken_commission_rate = obj.std_comm_rate
                                    commission_rate_diff = allowed_commission_rate - float(taken_commission_rate)
                                    cobl_amount = obj.fare_amount
                                    if cobl_amount:
                                        if commission_rate_diff < 0:
                                            # ADM
                                            adm_amount = (abs(commission_rate_diff) * cobl_amount) / 100
                                            comment = "Commission deducted " + str(
                                                taken_commission_rate) + "%. Carrier authorized " + str(
                                                allowed_commission_rate) + "%"
                                            allowed_commission_amount = (cobl_amount * allowed_commission_rate) / 100
                                            adms.append({
                                                'agency_no': obj.agency.agency_no,
                                                'trade_name': obj.agency.trade_name,
                                                'ticket_no': obj.ticket_no,
                                                'issue_date': obj.issue_date,
                                                'fare_amount': obj.fare_amount,
                                                'std_comm_amount': obj.std_comm_amount,
                                                'std_comm_rate': obj.std_comm_rate,
                                                'allowed_commission_amount': allowed_commission_amount,
                                                'amount': adm_amount,
                                                'comment': comment

                                            })

                                except Exception as e:
                                    pass

                        for obj in adm_objs:
                            if obj.transaction.agency.agency.agency_no not in exclude_list:
                                adms.append({
                                    'agency_no': obj.transaction.agency.agency_no,
                                    'trade_name': obj.transaction.agency.trade_name,
                                    'ticket_no': obj.transaction.ticket_no,
                                    'issue_date': obj.transaction.issue_date,
                                    'fare_amount': obj.transaction.fare_amount,
                                    'std_comm_amount': obj.transaction.std_comm_amount,
                                    'std_comm_rate': obj.transaction.std_comm_rate,
                                    'allowed_commission_amount': obj.allowed_commission_amount,
                                    'amount': obj.amount,
                                    'comment': obj.comment
                                })

            wb = xlwt.Workbook(style_compression=2)
            ws = wb.add_sheet('ADMS')

            # Sheet header, first row
            row_num = 0

            bold_center = xlwt.easyxf(
                "align: wrap yes, vert centre, horiz center;font: name Arial,height 280, bold True")
            wrap_data = xlwt.easyxf("align: wrap yes, vert centre, horiz center;font: name Arial,height 180")

            ws.row(row_num).height = 20 * 20
            ws.write_merge(row_num, 0, 0, 10, airline_name.upper(), bold_center)
            row_num = row_num + 1
            ws.row(row_num).height = 20 * 20
            ws.write_merge(row_num, 1, 0, 10, "ADM Reports".upper(), bold_center)
            row_num = row_num + 1
            ws.row(row_num).height = 20 * 20
            ws.write_merge(row_num, 2, 0, 10, month_year.upper(), bold_center)
            row_num = row_num + 1
            ws.row(row_num).height_mismatch = True
            ws.row(row_num).height = 20 * 36

            ws.col(0).width = (256 * 13)
            ws.col(1).width = (256 * 35)
            ws.col(2).width = (256 * 12)
            ws.col(3).width = (256 * 12)
            ws.col(4).width = (256 * 11)
            ws.col(5).width = (256 * 12)
            ws.col(6).width = (256 * 13)
            ws.col(7).width = (256 * 14)
            ws.col(8).width = (256 * 11)
            ws.col(9).width = (256 * 40)
            ws.col(10).width = (256 * 9)

            columns = ["Agency Number", "Agency Name", "Tkt Number", "Issue Date", "Tkt Base Fare",
                       "Deducted Commission",
                       "Commission Charged( %)", "Authorized Commission " + str(allowed_commission_rate) + " %",
                       "ADM Amount", "Comments", "ADM NO"]

            for col_num in range(len(columns)):
                ws.write(row_num, col_num, columns[col_num], xlwt.easyxf(
                    "align: wrap yes, vert centre, horiz center;font: name Arial, bold True, height 180;pattern: pattern solid,fore-colour yellow;border: left thin,right thin,top thin,bottom thin"))


            for d in adms:
                row_num = row_num + 1
                ws.row(row_num).height = 256 * 3
                col_num = 0
                ws.write(row_num, col_num, d.get('agency_no'), wrap_data)
                col_num = col_num + 1
                ws.write(row_num, col_num, d.get('trade_name'), wrap_data)
                col_num = col_num + 1
                ws.write(row_num, col_num, d.get('ticket_no'), wrap_data)
                col_num = col_num + 1
                ws.write(row_num, col_num, d.get('issue_date').strftime("%Y-%m-%d"), wrap_data)
                col_num = col_num + 1
                ws.write(row_num, col_num, d.get('fare_amount'), wrap_data)
                col_num = col_num + 1
                ws.write(row_num, col_num, d.get('std_comm_amount'), wrap_data)
                col_num = col_num + 1
                ws.write(row_num, col_num, d.get('std_comm_rate'), wrap_data)
                col_num = col_num + 1
                ws.write(row_num, col_num, d.get('allowed_commission_amount'), wrap_data)
                col_num = col_num + 1
                ws.write(row_num, col_num, d.get('amount'), wrap_data)
                col_num = col_num + 1
                ws.write(row_num, col_num, d.get('comment'), wrap_data)
                col_num = col_num + 1
                ws.write(row_num, col_num, '', wrap_data)
                col_num = col_num + 1

            wb.save(response)
            return response


class SalesSummaryReport(PermissionRequiredMixin, TemplateView):
    """SalesSummaryReport."""

    permission_required = ('report.view_sales_summary',)
    template_name = 'sales-summary.html'

    def get_context_data(self, **kwargs):
        # from asplinks.celery import download_files
        # download_files(manual=True)


###########################################################
        # from main.models import Country
        from datetime import date
        # import logging
        # from report.models import Transaction, Disbursement
        # day = date.today()
        # one_day = timedelta(1)
        # start_date = 0
        # while True:
        #     x = Disbursement.objects.filter(
        #         report_period__ped=day)
        #     if len(x):
        #         break
        #     day = day - one_day
        # for country in Country.objects.all():
        #     airlines = country.airlines.all()
        #     transactions = Transaction.objects.filter(
        #         report__airline__in=airlines,
        #         report__report_period__ped=day).values_list('agency__agency_no',flat=True)
        #     logging.info(str(transactions))

        def new_agent_report_generate():
            from main.models import Airline
            from report.models import Transaction, Disbursement, NewAgents
            day = date.today()
            one_day = timedelta(1)
            while True:
                x = Disbursement.objects.filter(
                    report_period__ped=day)
                if len(x):
                    break
                day = day - one_day
            for airline in Airline.objects.filter(country__code='US'):
                transactions = Transaction.objects.filter(report__airline=airline)
                transaction_fltrd = transactions.filter(
                    report__report_period__ped=day)
                for transaction in transaction_fltrd:
                    other_trans = transactions.filter(
                        agency=transaction.agency)
                    if len(other_trans) == 1:
                        NewAgents.objects.get_or_create(agency=transaction.agency,airline=airline,ped=day)

        # new_agent_report_generate()

#############################################################

        context = super(SalesSummaryReport, self).get_context_data(**kwargs)
        month_year = self.request.GET.get('month_year', '')
        airline = self.request.GET.get('airline', '')

        country = self.request.session.get('country')
        is_arc_var = is_arc(self.request.session.get('country'))
        if airline and month_year:
            # Transaction.objects.filter(report__airline__name = "AIR ASTANA",report__country__name = "United States", report__report_period__year= 2023,report__report_period__month= 3).delete()
            # Transaction.objects.filter(report__airline__name = "AIRLINK",report__country__name = "United States", report__report_period__year= 2023,report__report_period__month= 3).delete()
            month = datetime.datetime.strptime(month_year, '%B %Y').month or ''
            year = datetime.datetime.strptime(month_year, '%B %Y').year or ''

            airline_data = Airline.objects.get(id=airline)

            iata_coordination_fee = 0.00
            gsa_commision = 0.00
            sub_tax_yq_yr = Charges.objects.select_related('transaction').filter(type__in=['YQ', 'YR'],transaction__report__report_period__ped=OuterRef('report__report_period__ped'),transaction__report__airline=airline).values('transaction__report__report_period__ped').annotate(
                dcount=Count('transaction__report__report_period__ped'),total_tax_yq_yr=Sum('amount', output_field=FloatField())).values('total_tax_yq_yr')

            sub_tax_cp_mf = Charges.objects.select_related('transaction').filter(type__in=['CP', 'MF'],
                                                                                 transaction__report__report_period__ped=OuterRef(
                                                                                     'report__report_period__ped'),
                                                                                 transaction__report__airline=airline).values(
                'transaction__report__report_period__ped').annotate(
                dcount=Count('transaction__report__report_period__ped'),
                total_tax_cp_mf=Sum('amount', output_field=FloatField())).values('total_tax_cp_mf')





            sub_tax = Taxes.objects.select_related('transaction').filter(
                transaction__report__report_period__ped=OuterRef('report__report_period__ped'),
                transaction__report__airline=airline).values(
                'transaction__report__report_period__ped').annotate(
                dcount=Count('transaction__report__report_period__ped'),
                total_tax=Sum('amount', output_field=FloatField())).values('total_tax')

            total_adm = AgencyDebitMemo.objects.filter(transaction__report__airline=airline,
                                                       transaction__report__report_period__ped=OuterRef(
                                                           'report__report_period__ped')).values(
                'transaction__report__report_period__ped').annotate(
                dcount=Count('transaction__report__report_period__ped'),
                total_adm=Sum('amount', output_field=FloatField())).values('total_adm')
            if is_arc_var:

                month_date = datetime.datetime.strptime(month_year, '%B %Y')

                commission_history = CommissionHistory.objects.filter(airline=airline, type='A',
                                                                      from_date__lte=month_date)
                if commission_history.exists():
                    iata_coordination_fee = commission_history.order_by('-from_date').first().rate

                commission_history = CommissionHistory.objects.filter(airline=airline, type='D',
                                                                      from_date__lte=month_date)

                if commission_history.exists():
                    gsa_commision = commission_history.order_by('-from_date').first().rate


                transactions = Transaction.objects.exclude(transaction_type__startswith='SP',
                                                           agency__agency_no='6999001').select_related('report',
                                                                                                       'report__airline',
                                                                                                       'report__report_period').filter(
                    report__airline=airline,
                    report__report_period__month=month,
                    report__report_period__year=year).values(
                    'report__report_period__ped', 'report__report_period__week', 'card_type',
                    'report__report_period__remittance_date', 'report__cc', 'report__ca').annotate(
                    total_fare_amount=Sum('fare_amount'),
                    total_commission=Sum('std_comm_amount'),
                    pen_value=Sum('pen'),
                    cp=Sum('pen', filter=Q(report__airline=airline, pen_type='CANCEL PEN')),
                    total_ap_acm=Sum('transaction_amount',
                                     filter=Q(report__airline=airline, transaction_type__startswith='ACM',
                                              agency__agency_no='31768026')),

                ).order_by(
                    'report__report_period__ped')


                transactions_headers = []
                transactions_rows = []
                transactions_rows_iata = []
                transactions_rows_gsa = []
                total_due_to_airlinepros = []
                total_due_to_airlinepros_negative = []
                acms_rows = []
                acms_rows_total = []

                total_ap_acm = 0.00
                weekly_sales_total = 0.00
                total_cc = 0.00
                total_ca = 0.00

                net_sales_weekly_total = 0.00
                net_sales_weekly_total_iata = 0.00
                net_sales_weekly_total_gsa = 0.00
                pending_deduction_total = 0.00
                weekly_total_cash_disbursement_total = 0.00
                arc_deductions_total = 0.00
                cash_amount = 0
                test = 0
                add_yqyr=False



                qset= CommissionHistory.objects.filter(airline = airline).first()

                if qset:
                   add_yqyr=qset.add_yq_yr


                for t in transactions:

                    transactions_headers.append(t.get('report__report_period__ped'))

                    total_tax_yq_yr = Charges.objects.select_related('transaction').filter(type__in=['YQ', 'YR'],
                                                                                           transaction__report__report_period__ped=t.get(
                                                                                               'report__report_period__ped'),
                                                                                           transaction__report__airline=airline).aggregate(
                        total_tax_yq_yr=Coalesce(Sum('amount'), V(0))).get('total_tax_yq_yr')



                    total_tax_cp_mf = Charges.objects.select_related('transaction').filter(type__in=['CP', 'MF'],
                                                                                           transaction__report__report_period__ped=t.get(
                                                                                               'report__report_period__ped'),
                                                                                           transaction__report__airline=airline).aggregate(
                        total_tax_cp_mf=Coalesce(Sum('amount'), V(0))).get('total_tax_cp_mf')

                    penalty_charge = Charges.objects.select_related('transaction').filter(type__in=['CP'],
                                                                                           transaction__report__report_period__ped=t.get(
                                                                                               'report__report_period__ped'),
                                                                                           transaction__report__airline=airline).aggregate(
                        total_tax_cp_mf=Coalesce(Sum('amount'), V(0))).get('total_tax_cp_mf')

                    # pen_chrg_2 = Charges.objects.select_related('transaction')
                    # print("penalty charge",penalty_charge)


                    if t.get('total_fare_amount'):

                        total_fare_amount_var = t.get('total_fare_amount')
                    else:
                        total_fare_amount_var = 0.00

                    if t.get('cp'):
                        cp_var = t.get('cp')
                    else:
                        cp_var = 0.00

                    # total_fare_amount_var = total_fare_amount_var + cp_var
                    total_fare_amount_var = total_fare_amount_var + penalty_charge

                    total_ap_acm_var = 0.00
                    if t.get('total_ap_acm'):
                        total_ap_acm_var = t.get('total_ap_acm')

                    total_fare_amount_var = total_fare_amount_var - total_ap_acm_var
                    net_sales_weekly = total_fare_amount_var

                    total_tax_var = Taxes.objects.select_related('transaction').filter(
                        transaction__report__report_period__ped=t.get('report__report_period__ped'),
                        transaction__report__airline=airline).aggregate(total_tax=Coalesce(Sum('amount'), V(0))).get(
                        'total_tax')

                    if t.get('total_commission'):
                        total_commission_var = t.get('total_commission')*-1
                        # total_commission_var = abs(t.get('total_commission')) * -1
                    else:
                        total_commission_var = 0.00
                    weekly_sales = total_fare_amount_var + total_tax_var + total_commission_var + total_tax_yq_yr + value_check(
                        t.get('total_ap_acm'))

                    test += total_commission_var
                    remittance = t.get('report__report_period__remittance_date')


                    try:
                        disb = Disbursement.objects.get(report_period__ped=t.get('report__report_period__ped'),
                                                        airline=airline)
                        cash_disbursement = disb.bank7
                        arc_deductions = disb.arc_deduction + disb.arc_fees + disb.arc_reversal + disb.arc_tot
                        weekly_total_cash_disbursement = disb.arc_net_disb
                    except Exception as e:
                        cash_disbursement = 0.00
                        arc_deductions = 0.00
                        weekly_total_cash_disbursement = 0.00

                    try:
                        pending_deduction = Deduction.objects.filter(
                            report__report_period__ped=t.get('report__report_period__ped'), report__airline=airline,
                            pending=True).aggregate(pending_deduction=Coalesce(Sum('amount'), V(0))).get(
                            'pending_deduction')
                    except Exception as e:
                        pending_deduction = 0.00

                    arc_deductions_total = arc_deductions_total + arc_deductions
                    weekly_total_cash_disbursement_total = weekly_total_cash_disbursement_total + weekly_total_cash_disbursement
                    card_disbursement = weekly_sales - cash_disbursement

                    pending_deduction_total = pending_deduction_total + pending_deduction


                    ped_total = total_fare_amount_var
                    pen_val = t.get("pen_value")
                    if t.get("pen_value") is None:
                        pen_val = 0
                    if not total_tax_cp_mf:
                        total_tax_cp_mf = 0

                    calculated_fare_amount = total_fare_amount_var + total_tax_cp_mf - total_commission_var

                    transactions_rows.append(
                        {"fare": total_fare_amount_var, "tax": total_tax_var, "comm": total_commission_var,
                         "calculated_fare_amount": calculated_fare_amount,
                         "pen": t.get("pen_value"), "tax_yq_yr": total_tax_yq_yr, "tax_cp_mf": total_tax_cp_mf,
                         "weekly_sales_total": weekly_sales,
                         "total_ca": cash_disbursement, "total_cc": card_disbursement,
                         "total_ap_acm": t.get("total_ap_acm"), "remittance": remittance,
                         "week": t.get('report__report_period__week'), "pending_deduction": pending_deduction,
                         'net_sales_weekly': net_sales_weekly, 'ped_total': ped_total, 'arc_deductions': arc_deductions,
                         'weekly_total_cash_disbursement': weekly_total_cash_disbursement})

                    net_sales_weekly_total = net_sales_weekly_total + net_sales_weekly

                    if t.get('total_ap_acm'):
                        total_ap_acm = total_ap_acm + t.get('total_ap_acm')

                    weekly_sales_total = weekly_sales_total + weekly_sales

                    if card_disbursement:
                        total_cc = total_cc + card_disbursement
                    if cash_disbursement:
                        total_ca = total_ca + cash_disbursement

                  #ifelse for add yqyr
                    if add_yqyr :
                       if  iata_coordination_fee !=0:

                              transactions_rows_iata.append(((net_sales_weekly * iata_coordination_fee) / 100)+total_tax_yq_yr)
                              net_sales_weekly_total_iata = net_sales_weekly_total_iata + ((
                        (net_sales_weekly * iata_coordination_fee) / 100)+total_tax_yq_yr)
                       else:
                           transactions_rows_iata.append((net_sales_weekly * iata_coordination_fee) / 100)

                       if gsa_commision !=0:

                        transactions_rows_gsa.append(((net_sales_weekly * gsa_commision) / 100) + total_tax_yq_yr )
                        # print("------------------trgsa-----------",transactions_rows_gsa)
                       else:
                           transactions_rows_gsa.append((net_sales_weekly * gsa_commision) / 100)


                       net_sales_weekly_total_gsa = net_sales_weekly_total_gsa + (
                        (net_sales_weekly * gsa_commision) / 100)

                       total_due_to_airlinepros.append(
                            ((net_sales_weekly * iata_coordination_fee) / 100) + ((net_sales_weekly * gsa_commision) / 100))
                       total_due_to_airlinepros_negative.append(((((net_sales_weekly * iata_coordination_fee) / 100) + (
                            (net_sales_weekly * gsa_commision) / 100))) * -1)

                 #ifelse for yqyr

                    else:
                        # ifelse iata
                        transactions_rows_iata.append((net_sales_weekly * iata_coordination_fee) / 100)
                        transactions_rows_gsa.append((net_sales_weekly * gsa_commision) / 100)
                        # print("transactions-------------iata=",transactions_rows_iata)


                        net_sales_weekly_total_iata = net_sales_weekly_total_iata + (
                            (net_sales_weekly * iata_coordination_fee) / 100)
                        net_sales_weekly_total_gsa = net_sales_weekly_total_gsa + (
                            (net_sales_weekly * gsa_commision) / 100)

                        total_due_to_airlinepros.append(
                            ((net_sales_weekly * iata_coordination_fee) / 100) + ((net_sales_weekly * gsa_commision) / 100))
                        total_due_to_airlinepros_negative.append(((((net_sales_weekly * iata_coordination_fee) / 100) + (
                            (net_sales_weekly * gsa_commision) / 100))) * -1)


            else:

                month_date = datetime.datetime.strptime(month_year, '%B %Y')

                commission_history = CommissionHistory.objects.filter(airline=airline, type='I',
                                                                  from_date__lte=month_date)





                if commission_history.exists():
                    iata_coordination_fee = commission_history.order_by('-from_date').first().rate

                commission_history = CommissionHistory.objects.filter(airline=airline, type='G',
                                                                      from_date__lte=month_date)
                if commission_history.exists():

                    gsa_commision = commission_history.order_by('-from_date').first().rate


                transactions = Transaction.objects.exclude(transaction_type__startswith='SP',
                                                           agency__agency_no='6999001').select_related('report',
                                                                                                       'report__airline',
                                                                                                       'report__report_period').filter(
                    report__airline=airline,
                    report__report_period__month=month,
                    report__report_period__year=year).values(
                    'report__report_period__ped', 'report__report_period__week',
                    'report__report_period__remittance_date').annotate(
                    dcount=Count('report__report_period__ped'),
                    total_fare_amount=Sum('fare_amount'),
                    total_commission=Sum('std_comm_amount'),
                    report__cc=Sum('cc'),
                    report__ca=Sum('ca'),
                    pen_value=Sum('pen'),
                    cp=Sum('pen', filter=Q(report__airline=airline, pen_type='CP')),
                    total_tax=Subquery(sub_tax, output_field=FloatField()),

                    total_tax_yq_yr=Subquery(sub_tax_yq_yr,
                                             output_field=FloatField()),
                    total_tax_cp_mf=Subquery(sub_tax_cp_mf, output_field=FloatField())
                ).order_by(
                    'report__report_period__ped')

                # transactions=transactions.exclude(transaction_type__startswith='ADM')
                # for star in transactions:
                #     print('001**',star.get("report__ca"))
                #     print('0022**',star.get("total_tax_yq_yr"))

                transactions_headers = []
                transactions_rows = []
                transactions_rows_iata = []
                transactions_rows_gsa = []
                total_due_to_airlinepros = []
                total_due_to_airlinepros_negative = []
                acms_rows = []
                acms_rows_total = []

                total_ap_acm = 0.00
                weekly_sales_total = 0.00
                total_cc = 0.00
                total_ca = 0.00

                net_sales_weekly_total = 0.00
                net_sales_weekly_total_iata = 0.00
                net_sales_weekly_total_gsa = 0.00
                pending_deduction_total = 0.00
                weekly_total_cash_disbursement_total = 0.00
                arc_deductions_total = 0.00
                total_ap_acm_var = 0.00


                if not transactions:

                    for c_t in range(0, 4):
                        t = {}
                        ap_acm = {}
                        month_day = self.get_all_sundays(month, year)[c_t]
                        transactions_headers.append(month_day)

                        total_tax_yq_yr = 0.00
                        total_tax_cp_mf = 0.00
                        total_fare_amount_var = 0.00
                        total_commission_var = 0.00
                        total_tax_var = 0.00
                        ap_acm_comm = 0.00
                        total_commission_var = 0.00
                        cp_var = 0.00
                        ap_acm_cp = 0.00
                        cp_var = 0.00


                        remittance = month_day

                        cash_disbursement = 0.00

                        pending_deduction_total = 0.00

                        net_sales_weekly = 0.00

                        weekly_sales = 0.00
                        pen_val = 0
                        if t.get("pen_value"):
                            pen_val = float(t.get("pen_value"))
                        calculated_fare_amount = "{:.2f}".format(
                            total_fare_amount_var + total_tax_cp_mf - total_commission_var)

                        transactions_rows.append(
                            {"fare": total_fare_amount_var, "tax": t.get('total_tax'),
                             "calculated_fare_amount": calculated_fare_amount,
                             "pen": t.get("pen_value"), "comm": (abs(total_commission_var) * -1),
                             "tax_yq_yr": total_tax_yq_yr, "tax_cp_mf": total_tax_cp_mf,
                             "weekly_sales_total": weekly_sales,
                             "total_ca": cash_disbursement, "total_cc": t.get("report__cc"),
                             "total_ap_acm": total_ap_acm_var, "remittance": remittance,
                             "week": c_t + 1, "pending_deduction": 0.00,
                             'net_sales_weekly': net_sales_weekly, 'ped_total': total_fare_amount_var})


                        net_sales_weekly_total = net_sales_weekly_total + net_sales_weekly

                        total_ap_acm = 0.00

                        weekly_sales_total = 0.00

                        total_cc = 0.00
                        total_ca = 0.00

                        transactions_rows_iata.append((net_sales_weekly * iata_coordination_fee) / 100)
                        transactions_rows_gsa.append((net_sales_weekly * gsa_commision) / 100)

                        net_sales_weekly_total_iata = net_sales_weekly_total_iata + (
                            (net_sales_weekly * iata_coordination_fee) / 100)
                        net_sales_weekly_total_gsa = net_sales_weekly_total_gsa + (
                            (net_sales_weekly * gsa_commision) / 100)

                        total_due_to_airlinepros.append(
                            ((net_sales_weekly * iata_coordination_fee) / 100) + (
                                (net_sales_weekly * gsa_commision) / 100))
                        total_due_to_airlinepros_negative.append(
                            ((((net_sales_weekly * iata_coordination_fee) / 100) + (
                                (net_sales_weekly * gsa_commision) / 100))) * -1)

                c_t = 0
                for t in transactions:
                    transactions_headers.append(t.get('report__report_period__ped'))



                    if t.get("total_tax_yq_yr"):
                        total_tax_yq_yr = t.get("total_tax_yq_yr")

                    else:
                        total_tax_yq_yr = 0.00

                    if t.get("total_tax_cp_mf"):
                        total_tax_cp_mf = t.get("total_tax_cp_mf")
                    else:
                        total_tax_cp_mf = 0.00

                    ap_acm = Transaction.objects.filter(report__report_period__ped=t.get('report__report_period__ped'),
                                                        report__airline=airline, transaction_type__startswith='ACM',
                                                        agency__agency_no='6998051').aggregate(
                        total=Sum('fare_amount', output_field=FloatField()),
                        total_comm=Sum('std_comm_amount', output_field=FloatField()),
                        total_cp=Sum('pen', output_field=FloatField()))

                    total_cp = 0
                    for ac in Transaction.objects.filter(report__report_period__ped=t.get('report__report_period__ped'),
                                                         report__airline=airline, transaction_type__startswith='ACM',
                                                         agency__agency_no='6998051'):
                        for chrge in ac.charges_set.filter(type__in=['CP', 'MF']):
                            total_cp += int(chrge.amount)

                    if t.get('total_fare_amount'):
                        total_ap_acm_var = 0
                        if ap_acm.get('total'):
                            total_ap_acm_var = ap_acm.get('total')
                        total_fare_amount_var = t.get('total_fare_amount') + total_cp
                    else:
                        total_fare_amount_var = 0.00

                    if t.get('total_tax'):
                        total_tax_var = t.get('total_tax')
                    else:
                        total_tax_var = 0.00
                    total_commission_var = 0.00
                    if t.get('total_commission'):
                        total_commission_var = t.get('total_commission')

                    ap_acm_comm = 0.00
                    if ap_acm.get('total_comm'):
                        ap_acm_comm = ap_acm.get('total_comm')

                    total_commission_var = total_commission_var - ap_acm_comm

                    cp_var = 0.00
                    if t.get('cp'):
                        cp_var = t.get('cp')


                    ap_acm_cp = 0.00

                    if total_cp:
                        ap_acm_cp = total_cp
                    cp_var = cp_var - ap_acm_cp



                    remittance = t.get('report__report_period__remittance_date')

                    cash_disbursement = t.get("report__ca") - t.get('total_commission')

                    pending_deduction = value_check(t.get("report__ca")) + value_check(t.get('total_adm')) - (
                        value_check(t.get('total_commission'))) - value_check(t.get('total_ap_acm'))

                    pending_deduction_total = pending_deduction_total + pending_deduction

                    try:
                        country_obj = Country.objects.get(id=country)
                        if country_obj.name.lower() in ['canada', 'united states']:
                            net_sales_weekly = total_fare_amount_var  + total_tax_cp_mf
                            # net_sales_weekly = total_fare_amount_var
                        else:
                            net_sales_weekly = t.get("total_fare_amount")
                    except:
                        net_sales_weekly = total_fare_amount_var - total_commission_var + total_tax_cp_mf
                    weekly_sales = net_sales_weekly + total_tax_var + total_tax_yq_yr
                    # print(">>___yq___yr__>>>>>>>>",total_tax_yq_yr)

                    pen_val = t.get("pen_value")
                    if t.get("pen_value") is None:
                        pen_val = 0
                    calculated_fare_amount = total_fare_amount_var + total_tax_cp_mf - total_commission_var


                    transactions_rows.append(
                        {"fare": total_fare_amount_var, "tax": t.get('total_tax'),
                         "calculated_fare_amount": calculated_fare_amount,
                         "comm": (abs(total_commission_var) * -1), "pen": t.get("pen_value"),
                         "tax_yq_yr": total_tax_yq_yr, "tax_cp_mf": total_tax_cp_mf, "weekly_sales_total": weekly_sales,
                         "total_ca": cash_disbursement, "total_cc": t.get("report__cc"),
                         "total_ap_acm": total_ap_acm_var, "remittance": remittance,
                         "week": t.get('report__report_period__week'), "pending_deduction": pending_deduction,
                         'net_sales_weekly': net_sales_weekly, 'ped_total': total_fare_amount_var})

                    net_sales_weekly_total = net_sales_weekly_total + net_sales_weekly

                    if total_ap_acm_var:
                        total_ap_acm = total_ap_acm + total_ap_acm_var

                    weekly_sales_total = weekly_sales_total + weekly_sales

                    if t.get('report__cc'):
                        total_cc = total_cc + t.get('report__cc')
                    if cash_disbursement:
                        total_ca = total_ca + cash_disbursement


                    transactions_rows_iata.append((net_sales_weekly * iata_coordination_fee) / 100)
                    transactions_rows_gsa.append((net_sales_weekly * gsa_commision) / 100)

                    net_sales_weekly_total_iata = net_sales_weekly_total_iata + (
                        (net_sales_weekly * iata_coordination_fee) / 100)
                    net_sales_weekly_total_gsa = net_sales_weekly_total_gsa + (
                        (net_sales_weekly * gsa_commision) / 100)

                    total_due_to_airlinepros.append(
                        ((net_sales_weekly * iata_coordination_fee) / 100) + ((net_sales_weekly * gsa_commision) / 100))
                    total_due_to_airlinepros_negative.append(((((net_sales_weekly * iata_coordination_fee) / 100) + (
                        (net_sales_weekly * gsa_commision) / 100))) * -1)

                if transactions and len(transactions) < 4:
                    m_c = [1, 2, 3, 4]
                    for i in transactions:
                        try:
                            m_c.remove(i.get('report__report_period__week'))
                        except:
                            pass
                    c_t = m_c[0]

                    t = {}
                    ap_acm = {}

                    month_day = self.get_all_sundays(month, year)[c_t - 1]
                    transactions_headers.insert(c_t - 1, month_day)

                    total_tax_yq_yr = 0.00
                    total_fare_amount_var = 0.00
                    total_commission_var = 0.00
                    total_tax_var = 0.00
                    ap_acm_comm = 0.00
                    total_commission_var = 0.00
                    cp_var = 0.00
                    ap_acm_cp = 0.00
                    cp_var = 0.00


                    remittance = month_day

                    cash_disbursement = 0.00

                    pending_deduction_total = 0.00

                    net_sales_weekly = 0.00

                    weekly_sales = 0.00

                    transactions_rows.insert(c_t - 1,
                                             {"fare": total_fare_amount_var, "tax": t.get('total_tax'),
                                              "comm": (abs(total_commission_var) * -1),
                                              "tax_yq_yr": total_tax_yq_yr, "weekly_sales_total": weekly_sales,
                                              "total_ca": cash_disbursement, "total_cc": t.get("report__cc"),
                                              "total_ap_acm": total_ap_acm_var, "remittance": remittance,
                                              "week": c_t, "pending_deduction": 0.00,
                                              'net_sales_weekly': net_sales_weekly, 'ped_total': total_fare_amount_var})


                    net_sales_weekly_total = net_sales_weekly_total + net_sales_weekly

                    total_ap_acm = 0.00

                    weekly_sales_total = 0.00

                    total_cc = 0.00
                    total_ca = 0.00

                    transactions_rows_iata.insert(c_t - 1, (net_sales_weekly * iata_coordination_fee) / 100)
                    transactions_rows_gsa.insert(c_t - 1, (net_sales_weekly * gsa_commision) / 100)

                    net_sales_weekly_total_iata = net_sales_weekly_total_iata + (
                        (net_sales_weekly * iata_coordination_fee) / 100)
                    net_sales_weekly_total_gsa = net_sales_weekly_total_gsa + (
                        (net_sales_weekly * gsa_commision) / 100)

                    total_due_to_airlinepros.insert(c_t - 1,
                                                    ((net_sales_weekly * iata_coordination_fee) / 100) + (
                                                        (net_sales_weekly * gsa_commision) / 100))
                    total_due_to_airlinepros_negative.insert(c_t - 1,
                                                             ((((net_sales_weekly * iata_coordination_fee) / 100) + (
                                                                 (net_sales_weekly * gsa_commision) / 100))) * -1)

            months_weekend = ReportPeriod.objects.filter(year=year, month=month, country=country).values_list('ped',
                                                                                                              flat=True).order_by(
                'ped')
            file_uploaded_date = ReportFile.objects.select_related('airline', 'report_period').filter(airline=airline,
                                                                                                      report_period__month=month,
                                                                                                      report_period__year=year,
                                                                                                      country=country).values_list(
                'report_period__ped', flat=True).order_by('report_period__ped')

            try:
                first_ped = months_weekend[0]
                last_ped = months_weekend[len(months_weekend) - 1]
                start = first_ped - datetime.timedelta(days=5)
                end = last_ped + datetime.timedelta(days=1)
                days = [start + datetime.timedelta(day) for day in range((end - start).days + 1)]

                details_daily = DailyCreditCardFile.objects.filter(airline=airline, date__gte=start,
                                                                   date__lte=end).values_list('date', flat=True)

                missing_dates = set(months_weekend) - set(file_uploaded_date)

                missing_dates_daily = set(days) - set(details_daily)

                if is_arc_var:
                    not_uploadedfiles = []
                    for dis in Disbursement.objects.filter(airline=airline, filedate__gte=start,
                                                           filedate__lte=end):
                        if not dis.file1 or not dis.file2:
                            not_uploadedfiles.append(dis)
            except IndexError as e:
                not_uploadedfiles = []
                missing_dates = {}
                missing_dates_daily = {}
                messages.add_message(self.request, messages.WARNING,
                                     'Calendar file not uploaded for this period.')

            credit_file_dates = CarrierDeductions.objects.filter(airline=airline, filedate__month=month,
                                                                 filedate__year=year).values_list('filedate', flat=True)
            disbursement_files = Disbursement.objects.filter(airline=airline, filedate__month=month,
                                                             filedate__year=year).values_list('filedate', flat=True)

            if is_arc_var:
                missing_dates_credit = []
                if months_weekend.exists():
                    missing_dates_credit_val = months_weekend[0]
                    if missing_dates_credit_val not in credit_file_dates:
                        missing_dates_credit = [missing_dates_credit_val]
                missing_dates_disb = set(months_weekend) - set(disbursement_files)
                missing_dates_daily = []

            else:
                not_uploadedfiles = []
                missing_dates_credit = []
                missing_dates_disb = []

            context['transactions_headers'] = transactions_headers
            context['transactions_rows'] = transactions_rows
            context['missing_dates'] = sorted(missing_dates)
            context['missing_dates_credit'] = sorted(missing_dates_credit)
            context['missing_dates_disb'] = sorted(missing_dates_disb)
            context['not_uploadedfiles'] = not_uploadedfiles
            context['missing_dates_daily'] = sorted(missing_dates_daily)
            context['missing_dates_count'] = len(context['missing_dates']) + len(context['missing_dates_disb']) + len(
                context['missing_dates_credit']) + len(context['missing_dates_daily']) + len(context['not_uploadedfiles'])
            context['transactions_rows_iata'] = transactions_rows_iata
            context['transactions_rows_gsa'] = transactions_rows_gsa
            context['total_due_to_airlinepros'] = total_due_to_airlinepros
            context['weekly_total_cash_disbursement_total'] = weekly_total_cash_disbursement_total
            context['arc_deductions_total'] = arc_deductions_total
            context['total_due_to_airlinepros_negative'] = total_due_to_airlinepros_negative

            context['net_sales_weekly_total'] = round(net_sales_weekly_total, 2)

            context['total_ap_acm'] = total_ap_acm
            context['weekly_sales_total'] = weekly_sales_total
            context['total_cc'] = total_cc
            context['total_ca'] = total_ca
            context['net_sales_weekly_total_iata'] = net_sales_weekly_total_iata
            context['net_sales_weekly_total_gsa'] = net_sales_weekly_total_gsa
            context['sum_total_due_to_airlinepros'] = net_sales_weekly_total_gsa + net_sales_weekly_total_iata
            context['pending_deduction_total'] = pending_deduction_total
            context['iata_coordination_fee'] = iata_coordination_fee

            context['gsa_commision'] = gsa_commision
        context['month_year'] = month_year
        context['selected_airline'] = airline
        context['airlines'] = Airline.objects.filter(country=self.request.session.get('country'))
        return context

    def get_all_sundays(self, month, year):
        import calendar

        sundays = []
        cal = calendar.Calendar()

        for day in cal.itermonthdates(year, month):
            if day.weekday() == 6 and day.month == month:
                sundays.append(day)
        return sundays


class GetSalesSummaryReport(PermissionRequiredMixin, View):
    """Filtered SalesSummary Report download as CSV."""

    permission_required = ('report.download_sales_summary',)

    def get(self, request, *args, **kwargs):
        month_year = self.request.GET.get('month_year', '')
        airline = self.request.GET.get('airline', '')
        airline_name = ''
        is_arc_var = is_arc(self.request.session.get('country'))
        response = HttpResponse(content_type='application/vnd.ms-excel')
        country = self.request.session.get('country')
        if airline and month_year:
            
            month = datetime.datetime.strptime(month_year, '%B %Y').month or ''
            year = datetime.datetime.strptime(month_year, '%B %Y').year or ''
            airline_data = Airline.objects.get(id=airline)
            airline_name = airline_data.name
            country_obj = Country.objects.get(id=country)
            country_name = country_obj.name
            currency_code = country_obj.currency
            response[
                'Content-Disposition'] = 'inline; filename=' + airline_data.abrev + ' Sales Summary ' + month_year +"- {} ({})".format(country_name,currency_code)+ '.xls'

            iata_coordination_fee = 0.00
            gsa_commision = 0.00

            sub_tax_yq_yr = Charges.objects.select_related('transaction').filter(type__in=['YQ', 'YR'],
                                                                                 transaction__report__report_period__ped=OuterRef(
                                                                                     'report__report_period__ped'),
                                                                                 transaction__report__airline=airline).values(
                'transaction__report__report_period__ped').annotate(
                dcount=Count('transaction__report__report_period__ped'),
                total_tax_yq_yr=Sum('amount', output_field=FloatField())).values('total_tax_yq_yr')

            sub_tax_cp_mf = Charges.objects.select_related('transaction').filter(type__in=['CP', 'MF'],
                                                                                 transaction__report__report_period__ped=OuterRef(
                                                                                     'report__report_period__ped'),
                                                                                 transaction__report__airline=airline).values(

                'transaction__report__report_period__ped').annotate(
                dcount=Count('transaction__report__report_period__ped'),
                total_tax_cp_mf=Sum('amount', output_field=FloatField())).values('total_tax_cp_mf')
            sub_tax = Taxes.objects.select_related('transaction').filter(
                transaction__report__report_period__ped=OuterRef('report__report_period__ped'),
                transaction__report__airline=airline).values(
                'transaction__report__report_period__ped').annotate(
                dcount=Count('transaction__report__report_period__ped'),
                total_tax=Sum('amount', output_field=FloatField())).values('total_tax')

            total_adm = AgencyDebitMemo.objects.filter(transaction__report__airline=airline,
                                                       transaction__report__report_period__ped=OuterRef(
                                                           'report__report_period__ped')).values(
                'transaction__report__report_period__ped').annotate(
                dcount=Count('transaction__report__report_period__ped'),
                total_adm=Sum('amount', output_field=FloatField())).values('total_adm')
            if is_arc_var:

                month_date = datetime.datetime.strptime(month_year, '%B %Y')

                commission_history = CommissionHistory.objects.filter(airline=airline, type='A',
                                                                      from_date__lte=month_date)
                if commission_history.exists():
                    iata_coordination_fee = commission_history.order_by('-from_date').first().rate

                commission_history = CommissionHistory.objects.filter(airline=airline, type='D',
                                                                      from_date__lte=month_date)
                if commission_history.exists():
                    gsa_commision = commission_history.order_by('-from_date').first().rate

                transactions = Transaction.objects.exclude(transaction_type__startswith='SP',
                                                           agency__agency_no='6999001').select_related('report',
                                                                                                       'report__airline',
                                                                                                       'report__report_period').filter(
                    report__airline=airline,
                    report__report_period__month=month,
                    report__report_period__year=year).values(
                    'report__report_period__ped', 'report__report_period__week',
                    'report__report_period__remittance_date', 'report__cc', 'report__ca').annotate(
                    total_fare_amount=Sum('fare_amount'),
                    total_commission=Sum('std_comm_amount'),
                    cp=Sum('pen', filter=Q(report__airline=airline, pen_type='CANCEL PEN')),

                    total_ap_acm=Sum('transaction_amount',
                                     filter=Q(report__airline=airline, transaction_type__startswith='ACM',
                                              agency__agency_no='31768026')),

                ).order_by(
                    'report__report_period__ped')

                transactions_headers = []
                transactions_rows = []
                transactions_rows_iata = []
                transactions_rows_gsa = []
                total_due_to_airlinepros = []
                total_due_to_airlinepros_negative = []
                acms_rows = []
                acms_rows_total = []

                total_ap_acm = 0.00
                weekly_sales_total = 0.00
                total_cc = 0.00
                total_ca = 0.00

                net_sales_weekly_total = 0.00
                net_sales_weekly_total_iata = 0.00
                net_sales_weekly_total_gsa = 0.00
                pending_deduction_total = 0.00
                weekly_total_cash_disbursement_total = 0.00
                arc_deductions_total = 0.00
                total_commission_var_p = 0
                total_cp_total = 0
                total_fare_total = 0

                cash_t = []
                amex_t = []
                visa_t = []
                mstr_t = []
                other_t = []

                for t in transactions:
                    
                    transactions_headers.append(t.get('report__report_period__ped'))

                    total_tax_yq_yr = Charges.objects.select_related('transaction').filter(type__in=['YQ', 'YR'],
                                                                                           transaction__report__report_period__ped=t.get(
                                                                                               'report__report_period__ped'),
                                                                                           transaction__report__airline=airline).aggregate(
                        total_tax_yq_yr=Coalesce(Sum('amount'), V(0))).get('total_tax_yq_yr')

                    total_tax_cp_mf = Charges.objects.select_related('transaction').filter(type__in=['CP', 'MF'],
                                                                                           transaction__report__report_period__ped=t.get(
                                                                                               'report__report_period__ped'),
                                                                                           transaction__report__airline=airline).aggregate(
                        total_tax_cp_mf=Coalesce(Sum('amount'), V(0))).get('total_tax_cp_mf')



                    penalty_charge = Charges.objects.select_related('transaction').filter(type__in=['CP'],
                                                                                           transaction__report__report_period__ped=t.get(
                                                                                               'report__report_period__ped'),
                                                                                           transaction__report__airline=airline).aggregate(
                        total_tax_cp_mf=Coalesce(Sum('amount'), V(0))).get('total_tax_cp_mf')



                    if t.get('total_fare_amount'):

                        print(t.get('total_fare_amount'),"//////////")
                        total_fare_amount_var = t.get('total_fare_amount')
                        total_fare_total += total_fare_amount_var
                    else:
                        total_fare_amount_var = 0.00

                    if t.get('cp'):
                        cp_var = t.get('cp')
                    else:
                        cp_var = 0.00
                    total_cp_total += cp_var
                    # total_fare_amount_var = total_fare_amount_var + cp_var
                    total_fare_amount_var = total_fare_amount_var + penalty_charge

                    total_ap_acm_var = 0.00
                    if t.get('total_ap_acm'):
                        total_ap_acm_var = t.get('total_ap_acm')


                    total_fare_amount_var = total_fare_amount_var - total_ap_acm_var
                    net_sales_weekly = total_fare_amount_var

                    total_tax_var = Taxes.objects.select_related('transaction').filter(
                        transaction__report__report_period__ped=t.get('report__report_period__ped'),
                        transaction__report__airline=airline).aggregate(
                        total_tax=Coalesce(Sum('amount'), V(0))).get('total_tax')



                    if t.get('total_commission'):
                        total_commission_var = t.get('total_commission') * -1
                        # total_commission_var = abs(t.get('total_commission')) * -1
                    else:
                        total_commission_var = 0.00
                    total_commission_var_p += total_commission_var

                    # weekly_sales = total_fare_amount_var - total_commission_var

                    weekly_sales = total_fare_amount_var + total_tax_var + total_commission_var + total_tax_yq_yr + value_check(
                        t.get('total_ap_acm'))

                    remittance = t.get('report__report_period__remittance_date')

                    try:
                        disb = Disbursement.objects.get(report_period__ped=t.get('report__report_period__ped'),
                                                        airline=airline)
                        cash_disbursement = disb.bank7
                        arc_deductions = disb.arc_deduction + disb.arc_fees + disb.arc_reversal + disb.arc_tot
                        weekly_total_cash_disbursement = disb.arc_net_disb
                    except Exception as e:
                        cash_disbursement = 0.00
                        arc_deductions = 0.00
                        weekly_total_cash_disbursement = 0.00

                    try:
                        pending_deduction = Deduction.objects.filter(
                            report__report_period__ped=t.get('report__report_period__ped'), report__airline=airline,
                            pending=True).aggregate(pending_deduction=Coalesce(Sum('amount'), V(0))).get(
                            'pending_deduction')
                    except Exception as e:
                        pending_deduction = 0.00

                    arc_deductions_total = arc_deductions_total + arc_deductions

                    weekly_total_cash_disbursement_total = weekly_total_cash_disbursement_total + weekly_total_cash_disbursement
                    # print("weekly total --->",t.get('total_ap_acm'))
                    card_disbursement = weekly_sales - cash_disbursement


                    pending_deduction_total = pending_deduction_total + pending_deduction


                    ped_total = total_fare_amount_var
                    pen_val = 0
                    if t.get("pen_value"):
                        pen_val = float(t.get("pen_value"))
                    if not total_tax_cp_mf:
                        total_tax_cp_mf = 0
                    calculated_fare_amount = "{:.2f}".format(
                        total_fare_amount_var + total_tax_cp_mf - total_commission_var)
                    transactions_rows.append(
                        {"fare": total_fare_amount_var, "tax": total_tax_var, "comm": total_commission_var,
                         "calculated_fare_amount": calculated_fare_amount,
                         "tax_yq_yr": total_tax_yq_yr, "weekly_sales_total": weekly_sales,
                         "total_ca": cash_disbursement, "total_cc": card_disbursement,
                         "total_ap_acm": t.get("total_ap_acm"), "remittance": remittance,
                         "week": t.get('report__report_period__week'), "pending_deduction": pending_deduction,
                         'net_sales_weekly': net_sales_weekly, 'ped_total': ped_total, 'arc_deductions': arc_deductions,
                         'weekly_total_cash_disbursement': weekly_total_cash_disbursement})

                    net_sales_weekly_total = net_sales_weekly_total + net_sales_weekly

                    if t.get('total_ap_acm'):
                        total_ap_acm = total_ap_acm + t.get('total_ap_acm')

                    weekly_sales_total = weekly_sales_total + weekly_sales

                    if card_disbursement:
                        total_cc = total_cc + card_disbursement
                    if cash_disbursement:
                        total_ca = total_ca + cash_disbursement


                    transactions_rows_iata.append((net_sales_weekly * iata_coordination_fee) / 100)
                    transactions_rows_gsa.append((net_sales_weekly * gsa_commision) / 100)

                    net_sales_weekly_total_iata = net_sales_weekly_total_iata + (
                        (net_sales_weekly * iata_coordination_fee) / 100)
                    net_sales_weekly_total_gsa = net_sales_weekly_total_gsa + (
                        (net_sales_weekly * gsa_commision) / 100)

                    total_due_to_airlinepros.append(
                        ((net_sales_weekly * iata_coordination_fee) / 100) + ((net_sales_weekly * gsa_commision) / 100))
                    total_due_to_airlinepros_negative.append(((((net_sales_weekly * iata_coordination_fee) / 100) + (
                        (net_sales_weekly * gsa_commision) / 100))) * -1)
                

            else:

                month_date = datetime.datetime.strptime(month_year, '%B %Y')

                commission_history = CommissionHistory.objects.filter(airline=airline, type='I',
                                                                      from_date__lte=month_date)
                if commission_history.exists():
                    iata_coordination_fee = commission_history.order_by('-from_date').first().rate

                commission_history = CommissionHistory.objects.filter(airline=airline, type='G',
                                                                      from_date__lte=month_date)
                if commission_history.exists():
                    gsa_commision = commission_history.order_by('-from_date').first().rate

                transactions = Transaction.objects.exclude(transaction_type__startswith='SP',
                                                           agency__agency_no='6999001').select_related('report',
                                                                                                       'report__airline',
                                                                                                       'report__report_period').filter(
                    report__airline=airline,
                    report__report_period__month=month,
                    report__report_period__year=year).values(
                    'report__report_period__ped', 'report__report_period__week',
                    'report__report_period__remittance_date').annotate(
                    dcount=Count('report__report_period__ped'),
                    total_fare_amount=Sum('fare_amount'),
                    total_commission=Sum('std_comm_amount'),
                    report__cc=Sum('cc'),
                    report__ca=Sum('ca'),
                    cp=Sum('pen', filter=Q(report__airline=airline, pen_type='CP')),
                    total_tax=Subquery(sub_tax, output_field=FloatField()),
                    total_adm=Subquery(total_adm, output_field=FloatField()),

                    total_tax_yq_yr=Subquery(sub_tax_yq_yr,
                                             output_field=FloatField()),
                    total_tax_cp_mf=Subquery(sub_tax_cp_mf, output_field=FloatField())
                ).order_by(
                    'report__report_period__ped')

                transactions_headers = []
                transactions_rows = []
                transactions_rows_iata = []
                transactions_rows_gsa = []
                total_due_to_airlinepros = []
                total_due_to_airlinepros_negative = []
                acms_rows = []
                acms_rows_total = []

                total_ap_acm = 0.00
                weekly_sales_total = 0.00
                total_cc = 0.00
                total_ca = 0.00

                net_sales_weekly_total = 0.00
                net_sales_weekly_total_iata = 0.00
                net_sales_weekly_total_gsa = 0.00
                pending_deduction_total = 0.00
                weekly_total_cash_disbursement_total = 0.00
                arc_deductions_total = 0.00
                total_ap_acm_var = 0.00
                if not transactions:
                    for c_t in range(0, 4):
                        t = {}
                        ap_acm = {}
                        month_day = self.get_all_sundays(month, year)[c_t]
                        transactions_headers.append(month_day)

                        total_tax_yq_yr = 0.00
                        total_tax_cp_mf = 0.00
                        total_fare_amount_var = 0.00
                        total_commission_var = 0.00
                        total_tax_var = 0.00
                        ap_acm_comm = 0.00
                        total_commission_var = 0.00
                        cp_var = 0.00
                        ap_acm_cp = 0.00
                        cp_var = 0.00


                        remittance = month_day

                        cash_disbursement = 0.00

                        pending_deduction_total = 0.00

                        net_sales_weekly = 0.00

                        weekly_sales = 0.00
                        pen_val = 0
                        if t.get("pen_value"):
                            pen_val = float(t.get("pen_value"))
                        calculated_fare_amount = "{:.2f}".format(total_fare_amount_var + pen_val - total_commission_var)
                        transactions_rows.append(
                            {"fare": total_fare_amount_var, "tax": t.get('total_tax'),
                             "calculated_fare_amount": calculated_fare_amount,
                             "comm": (abs(total_commission_var) * -1),
                             "tax_yq_yr": total_tax_yq_yr, "tax_cp_mf": total_tax_cp_mf,
                             "weekly_sales_total": weekly_sales,
                             "total_ca": cash_disbursement, "total_cc": t.get("report__cc"),
                             "total_ap_acm": total_ap_acm_var, "remittance": remittance,
                             "week": c_t + 1, "pending_deduction": 0.00,
                             'net_sales_weekly': net_sales_weekly, 'ped_total': total_fare_amount_var})

                        net_sales_weekly_total = net_sales_weekly_total + net_sales_weekly

                        total_ap_acm = 0.00

                        weekly_sales_total = 0.00

                        total_cc = 0.00
                        total_ca = 0.00

                        transactions_rows_iata.append((net_sales_weekly * iata_coordination_fee) / 100)
                        transactions_rows_gsa.append((net_sales_weekly * gsa_commision) / 100)

                        net_sales_weekly_total_iata = net_sales_weekly_total_iata + (
                            (net_sales_weekly * iata_coordination_fee) / 100)
                        net_sales_weekly_total_gsa = net_sales_weekly_total_gsa + (
                            (net_sales_weekly * gsa_commision) / 100)

                        total_due_to_airlinepros.append(
                            ((net_sales_weekly * iata_coordination_fee) / 100) + (
                                (net_sales_weekly * gsa_commision) / 100))
                        total_due_to_airlinepros_negative.append(
                            ((((net_sales_weekly * iata_coordination_fee) / 100) + (
                                (net_sales_weekly * gsa_commision) / 100))) * -1)
                for t in transactions:
                    transactions_headers.append(t.get('report__report_period__ped'))

                    if t.get("total_tax_yq_yr"):
                        total_tax_yq_yr = t.get("total_tax_yq_yr")
                    else:
                        total_tax_yq_yr = 0.00

                    if t.get("total_tax_cp_mf"):
                        total_tax_cp_mf = t.get("total_tax_cp_mf")
                    else:
                        total_tax_cp_mf = 0.00

                    ap_acm = Transaction.objects.filter(report__report_period__ped=t.get('report__report_period__ped'),
                                                        report__airline=airline, transaction_type__startswith='ACM',
                                                        agency__agency_no='6998051').aggregate(
                        total=Sum('fare_amount', output_field=FloatField()),
                        total_comm=Sum('std_comm_amount', output_field=FloatField()),
                        total_cp=Sum('pen', output_field=FloatField()))

                    total_cp = 0
                    for ac in Transaction.objects.filter(report__report_period__ped=t.get('report__report_period__ped'),
                                                         report__airline=airline, transaction_type__startswith='ACM',
                                                         agency__agency_no='6998051'):
                        for chrge in ac.charges_set.filter(type__in=['CP', 'MF']):
                            total_cp += int(chrge.amount)

                    if t.get('total_fare_amount'):
                        total_ap_acm_var = 0.00
                        if ap_acm.get('total'):
                            total_ap_acm_var = ap_acm.get('total')
                        total_fare_amount_var = t.get('total_fare_amount') + total_cp
                    else:
                        total_fare_amount_var = 0.00
                    if t.get('total_tax'):
                        total_tax_var = t.get('total_tax')
                    else:
                        total_tax_var = 0.00

                    total_commission_var = 0.00
                    if t.get('total_commission'):
                        total_commission_var = t.get('total_commission')

                    ap_acm_comm = 0.00
                    if ap_acm.get('total_comm'):
                        ap_acm_comm = ap_acm.get('total_comm')

                    total_commission_var = total_commission_var - ap_acm_comm

                    cp_var = 0.00
                    if t.get('cp'):
                        cp_var = t.get('cp')

                    ap_acm_cp = 0.00
                    if ap_acm.get('total_cp'):
                        ap_acm_cp = ap_acm.get('total_cp')

                    cp_var = cp_var - ap_acm_cp


                    remittance = t.get('report__report_period__remittance_date')



                    cash_disbursement = t.get("report__ca") - t.get('total_commission')

                    pending_deduction = value_check(t.get("report__ca")) + value_check(t.get('total_adm')) - (
                        value_check(t.get('total_commission'))) - value_check(t.get('total_ap_acm'))

                    pending_deduction_total = pending_deduction_total + pending_deduction

                    try:
                        country_obj = Country.objects.get(id=country)
                        if country_obj.name.lower() in ['canada', 'united states']:
                            net_sales_weekly = total_fare_amount_var  + total_tax_cp_mf
                            # net_sales_weekly = total_fare_amount_var
                        else:
                            net_sales_weekly = t.get("total_fare_amount")
                    except:
                        net_sales_weekly = total_fare_amount_var - total_commission_var + total_tax_cp_mf


                    weekly_sales = net_sales_weekly + total_tax_var + total_tax_yq_yr
                    pen_val = 0
                    if t.get("pen_value"):
                        pen_val = float(t.get("pen_value"))
                    calculated_fare_amount = "{:.2f}".format(total_fare_amount_var + pen_val - total_commission_var)
                    transactions_rows.append(
                        {"fare": total_fare_amount_var, "tax": t.get('total_tax'),
                         "calculated_fare_amount": calculated_fare_amount,
                         "comm": (abs(total_commission_var) * -1),
                         "tax_yq_yr": total_tax_yq_yr, "tax_cp_mf": total_tax_cp_mf, "weekly_sales_total": weekly_sales,
                         "total_ca": cash_disbursement, "total_cc": t.get("report__cc"),
                         "total_ap_acm": total_ap_acm_var, "remittance": remittance,
                         "week": t.get('report__report_period__week'), "pending_deduction": pending_deduction,
                         'net_sales_weekly': net_sales_weekly, 'ped_total': total_fare_amount_var})

                    net_sales_weekly_total = net_sales_weekly_total + net_sales_weekly
                    if total_ap_acm_var:
                        total_ap_acm = total_ap_acm + total_ap_acm_var

                    weekly_sales_total = weekly_sales_total + weekly_sales

                    if t.get('report__cc'):
                        total_cc = total_cc + t.get('report__cc')
                    if cash_disbursement:
                        total_ca = total_ca + cash_disbursement



                    transactions_rows_iata.append((net_sales_weekly * iata_coordination_fee) / 100)
                    transactions_rows_gsa.append((net_sales_weekly * gsa_commision) / 100)

                    net_sales_weekly_total_iata = net_sales_weekly_total_iata + (
                        (net_sales_weekly * iata_coordination_fee) / 100)
                    net_sales_weekly_total_gsa = net_sales_weekly_total_gsa + (
                        (net_sales_weekly * gsa_commision) / 100)

                    total_due_to_airlinepros.append(
                        ((net_sales_weekly * iata_coordination_fee) / 100) + ((net_sales_weekly * gsa_commision) / 100))
                    total_due_to_airlinepros_negative.append(((((net_sales_weekly * iata_coordination_fee) / 100) + (
                        (net_sales_weekly * gsa_commision) / 100))) * -1)
            # sheet-start
            wb = xlwt.Workbook(style_compression=2)
            ws = wb.add_sheet('SalesSummary')

            # gsa_commision = 1.0

            ws.col(0).width = (256 * 12)
            # Sheet header, first row
            row_num = 0
            ssd_new=xlwt.easyxf("align: wrap yes, vert centre, horiz right;font: name Arial, height 180, bold True;border: left medium,right medium,top medium,bottom medium;pattern: pattern solid, fore_colour red")



            bold_center = xlwt.easyxf("font: name Arial, bold True, height 280; align: horiz center")
            wrap_data = xlwt.easyxf("align: wrap yes, vert centre, horiz center;font: name Arial,height 180")
            center_normal = xlwt.easyxf(
                "align: wrap yes, vert centre, horiz center;font: name Arial, height 180;border: left thin,right thin,top thin,bottom thin")
            center_normal_right = xlwt.easyxf(
                "align: wrap yes, vert centre, horiz right;font: name Arial, height 180;border: left thin,right thin,top thin,bottom thin")
            center_normal_border = xlwt.easyxf(
                "align: wrap yes, vert centre, horiz center;font: name Arial, height 180;border: left medium,right medium,top medium,bottom medium")

            center_normal_border_right = xlwt.easyxf(
                "align: wrap yes, vert centre, horiz right;font: name Arial, height 180, bold True;border: left medium,right medium,top medium,bottom medium")

            center_normal_border_yellow_center = xlwt.easyxf(
                "align: wrap yes, vert centre, horiz centre;font: name Arial, height 180, bold True;border: left medium,right medium,top medium,bottom medium;pattern: pattern solid, fore_colour yellow")



            center_normal_border_yellow_right = xlwt.easyxf(
                "align: wrap yes, vert centre, horiz right;font: name Arial, height 180, bold True;border: left medium,right medium,top medium,bottom medium;pattern: pattern solid, fore_colour yellow")
            center_normal_border_date = xlwt.easyxf(
                "align: wrap yes, vert centre, horiz center;font: name Arial, height 180, bold True;border: left medium,right medium,top medium,bottom medium")
            center_normal_border_color = xlwt.easyxf(
                "align: wrap yes, vert centre, horiz center;font: name Arial, height 180;border: left medium,right medium,top medium,bottom medium;pattern: pattern solid, fore_colour yellow")

            center_normal_border_color_right = xlwt.easyxf(
                "align: wrap yes, vert centre, horiz right;font: name Arial, height 180, bold True;border: left medium,right medium,top medium,bottom medium;pattern: pattern solid, fore_colour yellow")
            center_normal_color = xlwt.easyxf(
                "align: wrap yes, vert centre, horiz center;font: name Arial, height 180;pattern: pattern solid, fore_colour yellow;border: left thin,right thin,top thin,bottom thin")
            center_normal_color_right = xlwt.easyxf(
                "align: wrap yes, vert centre, horiz right;font: name Arial, height 180;pattern: pattern solid, fore_colour yellow;border: left thin,right thin,top thin,bottom thin")
            center_normal_red_right = xlwt.easyxf(
                "align: wrap yes, vert centre, horiz right;font: name Arial, height 180, colour red;pattern: pattern solid, fore_colour yellow;border: left thin,right thin,top thin,bottom thin")
            center_normal_color_date = xlwt.easyxf(
                "align: wrap yes, vert centre, horiz right;font: name Arial, height 180;pattern: pattern solid, fore_colour yellow;border: left thin,right thin,top thin,bottom thin")
            ws.row(0).height_mismatch = True
            ws.row(row_num).height = 20 * 22
            ws.write_merge(row_num, 0, 0, 10, airline_name.upper(), bold_center)
            row_num = row_num + 1
            ws.row(row_num).height = 20 * 22
            ws.write_merge(row_num, 1, 0, 10, "Sales Summary Reports".upper(), bold_center)
            row_num = row_num + 1
            ws.row(row_num).height = 20 * 22
            date_month_row = month_year.split(' ')[0][:3] + '-' + month_year.split(' ')[1]
            ws.write_merge(row_num, 2, 0, 10, date_month_row, bold_center)
            row_num = row_num + 1

            ws.col(0).width = (225 * 20)
            ws.write_merge(row_num, row_num + 1, 0, 0, "Week", center_normal_border)
            length = len(transactions_headers)
            count = 1
            ws.row(row_num + 1).height = 20 * 22

            for col_num, ele in enumerate(transactions_headers, 1):
                ws.write_merge(row_num, row_num, col_num * 4 - 3, col_num * 4, ele.strftime("%Y-%m-%d"),
                               center_normal_border_date)
                ws.col(count).width = (7 * 350)
                ws.write(row_num + 1, count, "Fare", center_normal_border)
                count = count + 1
                ws.col(count).width = (7 * 350)
                ws.write(row_num + 1, count, "Tax", center_normal_border)
                count = count + 1
                ws.col(count).width = (7 * 350)
                ws.write(row_num + 1, count, "YQ+YR", center_normal_border)
                count = count + 1
                ws.col(count).width = (7 * 350)
                ws.write(row_num + 1, count, "Comm", center_normal_border)
                count = count + 1

            ws.col(count).width = (8 * 350)
            ws.write_merge(row_num, row_num + 1, count, count, "AirlinePros ACM", center_normal_border_right )
            count = count + 1
            ws.col(count).width = (8 * 350)
            ws.write_merge(row_num, row_num + 1, count, count, "Weekly Sales Totals", center_normal_border_color)
            count = count + 1
            ws.col(count).width = (9 * 350)
            ws.write_merge(row_num, row_num + 1, count, count, "Weekly Cash Disbursement", center_normal_border_color)
            count = count + 1
            ws.col(count).width = (7 * 350)
            ws.write_merge(row_num, row_num + 1, count, count,
                           "IATA Deductions" if not is_arc_var else "ARC Deductions", center_normal_border_color)
            count = count + 1
            ws.col(count).width = (9 * 350)
            if is_arc_var:
                ws.write_merge(row_num, row_num + 1, count, count,
                               "Weekly total cash Disbursement", center_normal_border_color)
                count = count + 1
                ws.col(count).width = (9 * 350)
            ws.write_merge(row_num, row_num + 1, count, count, "Remittance" if not is_arc_var else "Disbursement Date",
                           center_normal_border_color)
            count = count + 1
            ws.col(count).width = (11 * 350)
            ws.write_merge(row_num, row_num + 1, count, count, "Weekly Credit Card Disbursement",
                           center_normal_border_color)
            count = count + 1
            ws.col(count).width = (8 * 350)

            ws.write_merge(row_num, row_num + 1, count, count, "Pending Deductions", center_normal_border_color)
            row_num = row_num + 2
            count = 1
            print(transactions_rows,"////////////")
            for num, ele in enumerate(transactions_rows):
                ws.write(row_num, 0, "Week" + str(ele.get("week")), center_normal)
                for i in range(1, count):
                    # empty cell for border
                    ws.write(row_num, i, '', xlwt.easyxf("border: left thin,right thin,top thin,bottom thin"))
                ws.write(row_num, count, format_value_excel(ele.get('fare')), center_normal_right)
                ws.write(row_num, count + 1, format_value_excel(ele.get('tax')), center_normal_right)
                ws.write(row_num, count + 2, format_value_excel(ele.get('tax_yq_yr')), center_normal_right)
                ws.write(row_num, count + 3, format_value_excel(ele.get('comm')), center_normal_right)
                for i in range(count + 4, length * 4 + 1):
                    # empty cell for border
                    ws.write(row_num, i, '', xlwt.easyxf("border: left thin,right thin,top thin,bottom thin"))



                col = length * 4

                ws.write(row_num, col + 1, format_value_excel(ele.get("total_ap_acm")), center_normal)
                ws.write(row_num, col + 2, format_value_excel(ele.get("weekly_sales_total")), center_normal_color_right)
                ws.write(row_num, col + 3, format_value_excel(ele.get("total_ca")), center_normal_color_right)
                ws.write(row_num, col + 4, format_value_excel(ele.get("arc_deductions")), center_normal_color_right)
                i = 4
                if is_arc_var:
                    i = i + 1
                    ws.write(row_num, col + i, format_value_excel(ele.get("weekly_total_cash_disbursement")),
                             center_normal_color_right)
                i = i + 1
                ws.write(row_num, col + i, ele.get("remittance").strftime("%Y-%m-%d") if ele.get("remittance") else '',
                         center_normal_color_date)
                i = i + 1
                ws.write(row_num, col + i, format_value_excel(ele.get("total_cc")), center_normal_color_right)
                print("--cc-----",format_value_excel(ele.get("total_cc")))
                i = i + 1
                ws.write(row_num, col + i, format_value_excel(ele.get("pending_deduction")), center_normal_red_right)
                row_num = row_num + 1
                count = count + 4

            # blank row
            row_num = row_num + 1
            count = 1
            ws.row(row_num).height = 20 * 22
            ws.write(row_num, 0, "Total", center_normal_border_yellow_right)
            ws.row(row_num + 2).height = 20 * 22
            ws.write(row_num + 2, 0, "Net sales per week", xlwt.easyxf(
                "align: wrap yes, vert centre, horiz left;font: name Arial, height 180, bold True;border: left thin,right thin,top thin,bottom thin"))
            ws.row(row_num + 3).height = 20 * 22
            ws.write(row_num + 3, 0, "IATA coordination fee (" + str(
                iata_coordination_fee) + " %)" if not is_arc_var else "ARC coordination fee (" + str(
                iata_coordination_fee) + " %)", xlwt.easyxf(
                "align: wrap yes, vert centre, horiz left;font: name Arial, height 180, bold True, colour red;border: left medium,right medium,top medium,bottom medium"))
            ws.row(row_num + 4).height = 20 * 22

            # --------------
            # If gsa commision > 0 , a new row is added in excel sheet. Otherwise the existing
            # row numbers are shifted one row up, so that there is no blank row in the sheet.
            # Since the columns belonging to GSA commision is written on different parts, this
            # condition is checked on all parts where GSA commision columns are written.


            if gsa_commision > 0:
                ws.write(row_num + 4, 0, "GSA commission (" + str(
                gsa_commision) + " %)" if not is_arc_var else "Distribution Intermediary Fee (" + str(
                gsa_commision) + " %)", xlwt.easyxf(
                "align: wrap yes, vert centre, horiz left;font: name Arial, height 180, bold True, colour red;border: left medium,right medium,top medium,bottom medium"))
            else:
                row_num-=1
            ws.row(row_num + 5).height = 20 * 22
            ws.write(row_num + 5, 0, "Total due to AirlinePros", xlwt.easyxf(
                "align: wrap yes, vert centre, horiz left;font: name Arial, height 180, bold True, colour red;pattern: pattern solid, fore_colour yellow;border: left medium,right medium,top medium,bottom medium"))
            if gsa_commision <= 0:
                row_num+=1

            for num, ele in enumerate(transactions_rows):
                #Total written in row_num = 8 --
                ws.write(row_num, count, format_value_excel(ele.get('ped_total')), center_normal_border_yellow_right)
                ws.write(row_num, count + 1, format_value_excel(ele.get('tax')), center_normal_border_yellow_right)
                ws.write(row_num, count + 2, format_value_excel(ele.get('tax_yq_yr')), center_normal_border_yellow_right)
                ws.write(row_num, count + 3, format_value_excel(ele.get('comm')), center_normal_border_yellow_right)
                #----
                ws.write(row_num + 2, count, format_value_excel(ele.get('net_sales_weekly')), center_normal_right)
                ws.write(row_num + 2, count + 1, '', xlwt.easyxf("border: left thin,right thin,top thin,bottom thin"))
                ws.write(row_num + 2, count + 2, '', xlwt.easyxf("border: left thin,right thin,top thin,bottom thin"))
                ws.write(row_num + 2, count + 3, '', xlwt.easyxf("border: left thin,right thin,top thin,bottom thin"))
                ws.write(row_num + 3, count, format_value_excel(transactions_rows_iata[num]), center_normal_right)
                ws.write(row_num + 3, count + 1, '', xlwt.easyxf("border: left thin,right thin,top thin,bottom thin"))
                ws.write(row_num + 3, count + 2, '', xlwt.easyxf("border: left thin,right thin,top thin,bottom thin"))
                ws.write(row_num + 3, count + 3, '', xlwt.easyxf("border: left thin,right thin,top thin,bottom thin"))
                if gsa_commision > 0:
                    ws.write(row_num + 4, count, format_value_excel(transactions_rows_gsa[num]), center_normal_right)
                    ws.write(row_num + 4, count + 1, '', xlwt.easyxf("border: left thin,right thin,top thin,bottom thin"))
                    ws.write(row_num + 4, count + 2, '', xlwt.easyxf("border: left thin,right thin,top thin,bottom thin"))
                    ws.write(row_num + 4, count + 3, '', xlwt.easyxf("border: left thin,right thin,top thin,bottom thin"))
                else:
                    row_num-=1
                ws.write(row_num + 5, count, format_value_excel(total_due_to_airlinepros[num]), xlwt.easyxf(
                    "align: wrap yes, vert centre, horiz right;font: name Arial, height 180, bold True;pattern: pattern solid, fore_colour yellow;border: left thin,right thin,top thin,bottom thin"))
                ws.write(row_num + 5, count + 1, "", xlwt.easyxf(
                    "pattern: pattern solid, fore_colour yellow;border: left thin,right thin,top thin,bottom thin"))
                ws.write(row_num + 5, count + 2, "", xlwt.easyxf(
                    "pattern: pattern solid, fore_colour yellow;border: left thin,right thin,top thin,bottom thin"))
                ws.write(row_num + 5, count + 3, "", xlwt.easyxf(
                    "pattern: pattern solid, fore_colour yellow;border: left thin,right thin,top thin,bottom thin"))
                if gsa_commision <= 0:
                    row_num+=1
                count = count + 4



            ws.write(row_num + 2, count, '', xlwt.easyxf("border: left thin,right thin,top thin,bottom thin"))
            ws.write(row_num + 2, count + 1, format_value_excel(net_sales_weekly_total), xlwt.easyxf(
                "align: wrap yes, vert centre, horiz right;font: name Arial, height 180, bold True;pattern: pattern solid, fore_colour yellow;border: left thin,right thin,top thin,bottom thin"))
            ws.write(row_num + 3, count, '', xlwt.easyxf("border: left thin,right thin,top thin,bottom thin"))
            ws.write(row_num + 3, count + 1, format_value_excel(net_sales_weekly_total_iata), xlwt.easyxf(
                "align: wrap yes, vert centre, horiz right;font: name Arial, height 180, bold True, colour red;pattern: pattern solid, fore_colour yellow;border: left thin,right thin,top thin,bottom thin"))

            if gsa_commision > 0:
                ws.write(row_num + 4, count, '', xlwt.easyxf("border: left thin,right thin,top thin,bottom thin"))
                ws.write(row_num + 4, count + 1, format_value_excel(net_sales_weekly_total_gsa), xlwt.easyxf(
                    "align: wrap yes, vert centre, horiz right;font: name Arial, height 180, bold True, colour red;pattern: pattern solid, fore_colour yellow;border: left thin,right thin,top thin,bottom thin"))
            else:
                row_num-=1

            ws.write(row_num + 5, count, "", xlwt.easyxf(
                "pattern: pattern solid, fore_colour yellow;border: left thin,right thin,top thin,bottom thin"))
            ws.write(row_num + 5, count + 1,
                     format_value_excel(net_sales_weekly_total_gsa + net_sales_weekly_total_iata),
                     xlwt.easyxf(
                         "align: wrap yes, vert centre, horiz right;font: name Arial, height 180, bold True;pattern: pattern solid, fore_colour yellow;border: left thin,right thin,top thin,bottom thin"))
            if gsa_commision <= 0:
                row_num+=1
            # empty cell for border
            for i in range(0, count):
                ws.write(row_num - 1, i, '', xlwt.easyxf("border: left thin,right thin,top thin,bottom thin"))

            ws.write(row_num - 1, count, "", xlwt.easyxf("border: left thin,right thin,top thin,bottom thin"))
            # changedd
            ws.write(row_num - 1, count + 1, "", center_normal_color)
            ws.write(row_num - 1, count + 2, "", center_normal_color)
            ws.write(row_num - 1, count + 3, "", center_normal_color)
            ws.write(row_num - 1, count + 4, "", center_normal_color)
            ws.write(row_num - 1, count + 5, "", center_normal_color)
            ws.write(row_num - 1, count + 6, "", center_normal_color)

            # print(">>>",row_num)
            # airline pro acm total written out in here.
            if is_arc_var:
                ws.write(row_num - 1, count + 7, "", center_normal_color)
            ws.write(row_num, count, format_value_excel(total_ap_acm),center_normal_border_yellow_center )
            #  changedd
            # print(">>>",row_num)
            # print(">>>",count)
            ws.write(row_num, count + 1, format_value_excel(weekly_sales_total), center_normal_border_color_right)
            ws.write(row_num, count + 2, format_value_excel(total_ca), center_normal_border_color_right)
            ws.write(row_num, count + 3, format_value_excel(arc_deductions_total), center_normal_border_color_right)

            i = 3
            if is_arc_var:
                i = i + 1
                ws.write(row_num, count + i, format_value_excel(weekly_total_cash_disbursement_total),
                         center_normal_border_color_right)
            i = i + 1
            ws.write(row_num, count + i, "", center_normal_border_color_right)
            i = i + 1
            ws.write(row_num, count + i, format_value_excel(total_cc), center_normal_border_color_right)
            i = i + 1
            ws.write(row_num, count + i, format_value_excel(pending_deduction_total), center_normal_border_color_right)
            count = 0
            row_num = row_num + 7
            if gsa_commision <= 0:
                row_num-=1
            ws.row(row_num).height = 20 * 22
            ws.write(row_num, 0, "ACMS SUBMITTED", xlwt.easyxf(
                "align: wrap yes, vert centre, horiz center;font: name Arial, height 180, bold True, colour red;pattern: pattern solid, fore_colour yellow;border: left medium,right medium,top medium,bottom medium"))
            row_num = row_num + 1
            ws.row(row_num).height = 20 * 22


            for num, ele in enumerate(total_due_to_airlinepros_negative):
                ws.write(row_num, 0, "ACM", xlwt.easyxf(
                    "align: wrap yes, vert centre, horiz left;font: name Arial, height 180, bold True, colour red;pattern: pattern solid, fore_colour yellow;border: left thin,right thin,top thin,bottom thin"))
                for i in range(1, count + 1):
                    # empty cell for border
                    ws.write(row_num, i, '', xlwt.easyxf("border: left thin,right thin,top thin,bottom thin"))

                ws.write(row_num, count + 1, format_value_excel(ele), xlwt.easyxf(
                    "align: wrap yes, vert centre, horiz right;font: name Arial, height 180, colour red;border: left thin,right thin,top thin,bottom thin"))

                for i in range(count + 2, length * 4 + 1):
                    # empty cell for border
                    ws.write(row_num, i, '', xlwt.easyxf("border: left thin,right thin,top thin,bottom thin"))
                col = length * 4
                ws.write(row_num, col + 1, '', xlwt.easyxf("border: left thin,right thin,top thin,bottom thin"))
                ws.write(row_num, col + 2, format_value_excel(ele), xlwt.easyxf(
                    "align: wrap yes, vert centre, horiz right;font: name Arial, height 180, colour red;pattern: pattern solid, fore_colour yellow;border: left thin,right thin,top thin,bottom thin"))

                row_num = row_num + 1
                count = count + 4
            ws.row(row_num).height = 20 * 22
            ws.write(row_num, 0, "ACM TOTALS PER PED", xlwt.easyxf(
                "align: wrap yes, vert centre, horiz left;font: name Arial, height 180, bold True, colour red;pattern: pattern solid, fore_colour yellow;border: left thin,right thin,top thin,bottom thin"))
            count = 1
            for num, ele in enumerate(total_due_to_airlinepros_negative):
                ws.write(row_num, count, format_value_excel(ele), xlwt.easyxf(
                    "align: wrap yes, vert centre, horiz right;font: name Arial, height 180, bold True, colour red;pattern: pattern solid, fore_colour yellow;border: left thin,right thin,top thin,bottom thin"))
                ws.write(row_num, count + 1, "", xlwt.easyxf(
                    "pattern: pattern solid, fore_colour yellow;border: left thin,right thin,top thin,bottom thin"))
                ws.write(row_num, count + 2, "", xlwt.easyxf(
                    "pattern: pattern solid, fore_colour yellow;border: left thin,right thin,top thin,bottom thin"))
                ws.write(row_num, count + 3, "", xlwt.easyxf(
                    "pattern: pattern solid, fore_colour yellow;border: left thin,right thin,top thin,bottom thin"))
                count = count + 4

            ws.write(row_num, count, "", xlwt.easyxf(
                "pattern: pattern solid, fore_colour yellow;border: left thin,right thin,top thin,bottom thin"))
            ws.write(row_num, count + 1,
                     format_value_excel((net_sales_weekly_total_gsa + net_sales_weekly_total_iata) * -1),
                     xlwt.easyxf(
                         "align: wrap yes, vert centre, horiz right;font: name Arial, height 180, bold True, colour red;pattern: pattern solid, fore_colour yellow;border: left thin,right thin,top thin,bottom thin"))

            wb.save(response)
        return response

    def get_all_sundays(self, month, year):
        import calendar

        sundays = []
        cal = calendar.Calendar()

        for day in cal.itermonthdates(year, month):
            if day.weekday() == 6 and day.month == month:
                sundays.append(day)
        return sundays


class TaxesPartial(PermissionRequiredMixin, View):
    """Taxes listing partial for ajax call."""

    permission_required = ('report.view_sales_details',)
    template_name = 'taxes-partial.html'

    def get(self, request, *args, **kwargs):
        context = dict()
        context['taxes'] = Taxes.objects.filter(transaction=self.kwargs['pk'])
        context['charges'] = Charges.objects.filter(transaction=self.kwargs['pk'])
        return render(request, self.template_name, context)


@task(time_limit=999999, soft_time_limit=999999)
def excel_sales_report(month_year, start_date, end_date, airline, is_arc_var, sales_version, to_email, base_url):
    qs = Transaction.objects.select_related('agency').prefetch_related('taxes_set').filter(is_sale=True).annotate(
        yq=Sum('charges__amount', filter=Q(charges__type='YQ')),
        yr=Sum('charges__amount', filter=Q(charges__type='YR')))
    if airline:
        qs = qs.filter(report__airline=airline)
    if month_year:
        month = datetime.datetime.strptime(month_year, '%B %Y').month or ''
        year = datetime.datetime.strptime(month_year, '%B %Y').year or ''
        if month and year:
            qs = qs.filter(report__report_period__month=month, report__report_period__year=year)

    if start_date and end_date:
        start = datetime.datetime.strptime(start_date, '%d %B %Y')
        end = datetime.datetime.strptime(end_date, '%d %B %Y')
        qs = qs.filter(report__report_period__ped__range=[start, end])

    airline_obj = Airline.objects.filter(pk=airline).first()
    airline_name = ''
    if airline_obj:
        airline_name = airline_obj.name
        if month_year:
            dt_rep = month_year
        else:
            dt_rep = start_date + " - " + end_date

        file_name = airline_obj.abrev + " " + dt_rep + " Sales Details.xls"
    else:
        file_name = "Sales Details.xls"

    wb = xlwt.Workbook(encoding='utf-8')
    ws = FitSheetWrapper(wb.add_sheet('Sales Details'))

    # Sheet header, first row
    row_num = 0

    font_style = xlwt.XFStyle()
    font_style.font.bold = True

    if airline_obj:
        ws.row(row_num).height = 20 * 20
        ws.write_merge(row_num, 0, 0, 10, airline_obj.name.upper(), bold_center)
        row_num = row_num + 1
        ws.row(row_num).height = 20 * 20
        ws.write_merge(row_num, 1, 0, 10, "SALES DETAILS REPORT", bold_center)
        row_num = row_num + 1
        ws.row(row_num).height = 20 * 20
        ws.write_merge(row_num, 2, 0, 10, dt_rep.upper(), bold_center)
        row_num = row_num + 1
        ws.row(row_num).height_mismatch = True
    else:
        ws.row(row_num).height = 20 * 20
        ws.write_merge(row_num, 0, 0, 10, "SALES DETAILS REPORT", bold_center)
        row_num = row_num + 1
        ws.row(row_num).height = 20 * 20
        ws.write_merge(row_num, 1, 0, 10, dt_rep.upper(), bold_center)
        row_num = row_num + 1
        ws.row(row_num).height_mismatch = True

    style_string = "font: bold on; pattern: pattern solid, fore_color yellow; borders: top_color black, bottom_color black, right_color black, left_color black,left thin, right thin, top thin, bottom thin;"
    yellow_style = xlwt.easyxf(style_string)

    default_columns = ["FOP", "Card Type" if not is_arc_var else "Card Number", "Agency", "Date", "Ticket", "Total",
                       "Fare amount" if not is_arc_var else "Base",
                       "Comm.Amt", "Comm. Rate(%)", "Trnc", "PEN", "Tax YQ", "Tax_YR"]

    if is_arc_var:
        del default_columns[0]

    peds = qs.values_list('report__report_period__ped', flat=True).order_by('report__report_period__ped').distinct()

    columns = ["FOP", "Card Type" if not is_arc_var else "Card Number", "Agency", "Date", "Ticket", "Total",
               "Fare amount" if not is_arc_var else "Base",
               "Comm.Amt", "Comm.Rate(%)", "Type", "Cancellation Penalty"]

    if is_arc_var:
        del columns[0]

    if not sales_version:
        columns = ["FOP", "Card Type" if not is_arc_var else "Card Number", "Agency", "Date", "Ticket", "Total",
                   "Fare amount" if not is_arc_var else "Base",
                   "Comm.Amt", "Comm.Rate(%)", "Type", "Cancellation Penalty", "Tax YQ", "Tax_YR"]
        if is_arc_var:
            del columns[0]
        taxes = set(qs.values_list("taxes__type", flat=True))
        for tx_type in taxes:
            if tx_type and tx_type not in ['YQ', 'YR']:
                columns.append("Tax " + tx_type)

    for ped in peds:
        if sales_version:
            transaction_type_counts = qs.filter(report__report_period__ped=ped).aggregate(
                tickets=Count('pk', filter=Q(transaction_type='TKTT') | Q(transaction_type='TKT')),
                refunds=Count('pk', filter=Q(transaction_type='RFND')),
                exchanges=Count('pk', filter=Q(fop='EX') | Q(transaction_type='EXCH')))
            ws.write(row_num, 2, "# Tickets = " + str(transaction_type_counts.get('tickets')), font_style)
            ws.write(row_num, 4, "# Refunds = " + str(transaction_type_counts.get('refunds')), font_style)
            ws.write(row_num, 6, "# Exchanges = " + str(transaction_type_counts.get('exchanges')), font_style)

        row_num = row_num + 1

        ws.write(row_num, 0, "PED: " + ped.strftime("%Y-%m-%d"), font_style)
        row_num = row_num + 1

        for col_num in range(len(columns)):
            ws.write(row_num, col_num, columns[col_num], yellow_style)

        for item in qs.filter(report__report_period__ped=ped):
            row_num = row_num + 1
            tax_start = len(default_columns)
            col_num = 0
            if not is_arc_var:
                ws.write(row_num, col_num, 'CA+CC' if item.ca and item.cc else (item.fop or ''))
                col_num = col_num + 1

            ws.write(row_num, col_num, item.card_type or item.card_code)
            col_num = col_num + 1
            ws.write(row_num, col_num, item.agency.agency_no or '')
            col_num = col_num + 1
            ws.write(row_num, col_num, item.issue_date.strftime("%Y-%m-%d") if item.issue_date else '')
            col_num = col_num + 1
            ws.write(row_num, col_num, item.ticket_no or '')
            col_num = col_num + 1
            ws.write(row_num, col_num, item.transaction_amount)
            col_num = col_num + 1
            ws.write(row_num, col_num, item.fare_amount or 0)
            col_num = col_num + 1
            ws.write(row_num, col_num, item.std_comm_amount or 0)
            col_num = col_num + 1
            ws.write(row_num, col_num, item.std_comm_rate or 0)
            col_num = col_num + 1
            ws.write(row_num, col_num, item.transaction_type or '')
            col_num = col_num + 1
            ws.write(row_num, col_num, item.pen or 0)
            col_num = col_num + 1

            if not sales_version:
                ws.write(row_num, col_num, item.yq or 0)
                col_num = col_num + 1
                ws.write(row_num, col_num, item.yr or 0)
                col_num = col_num + 1

                for tax in taxes:
                    if tax:
                        tax_amounts = [tx.amount for tx in item.taxes_set.all() if tx.type == tax]
                        if len(tax_amounts) > 0:
                            tax_amt = sum(tax_amounts)
                            ws.write(row_num, tax_start, tax_amt)
                        else:
                            ws.write(row_num, tax_start, 0)

                        tax_start = tax_start + 1
        row_num = row_num + 1
        ws.write(row_num, 0, "Total for report:", yellow_style)
        tot_for_rep = qs.filter(report__report_period__ped=ped).aggregate(total=Sum('transaction_amount')).get(
            'total') - qs.filter(report__report_period__ped=ped).aggregate(
            total=Sum('std_comm_amount')).get('total')
        if is_arc_var:
            ws.write(row_num, 4, format_value(tot_for_rep), yellow_style)
        else:
            ws.write(row_num, 5, format_value(tot_for_rep), yellow_style)
        row_num = row_num + 1
        ws.write(row_num, 0, "")
        row_num = row_num + 1
        path = os.path.join(settings.MEDIA_ROOT, 'excelreports', file_name)
        wb.save(path)

    excel = ExcelReportDownload()
    excel.file.name = 'excelreports/' + file_name
    excel.save()
    context = {'user': to_email,
               'link': base_url + excel.file.url,
               'report_name': 'Sales Detail',
               'airline_name': airline_name,
               'date': dt_rep
               }
    send_mail('Sales Detail ' + dt_rep + ': ' + airline_name, "email/excel-report-email.html", context,
              [to_email], from_email='assda@assda.com')
    return True


class GetSalesReport(PermissionRequiredMixin, View):
    """Filtered Sales Report download as exel."""

    permission_required = ('report.download_sales_details',)

    def get(self, request, *args, **kwargs):
        month_year = self.request.GET.get('month_year', '')
        start_date = self.request.GET.get('start_date', '')
        end_date = self.request.GET.get('end_date', '')
        airline = self.request.GET.get('airline', '')
        is_arc_var = is_arc(self.request.session.get('country'))
        sales_version = self.request.GET.get('sales_version')



        qs = Transaction.objects.select_related('agency').prefetch_related('taxes_set').exclude(
            transaction_type__startswith='ACM',
            agency__agency_no__in=['6998051', ]).annotate(
            yq=Sum('charges__amount', filter=Q(charges__type='YQ')),
            yr=Sum('charges__amount', filter=Q(charges__type='YR')))




        if airline:
            qs = qs.filter(report__airline=airline)
        if month_year:
            month = datetime.datetime.strptime(month_year, '%B %Y').month or ''
            year = datetime.datetime.strptime(month_year, '%B %Y').year or ''
            if month and year:
                qs = qs.filter(report__report_period__month=month, report__report_period__year=year)
                #qs= qs.filter(issue_date__year= year).filter(issue_date__month= month).order_by('issue_date')


        if start_date and end_date:
            start = datetime.datetime.strptime(start_date, '%d %B %Y')
            end = datetime.datetime.strptime(end_date, '%d %B %Y')
            qs = qs.filter(report__report_period__ped__range=[start, end])
            #qs = qs.filter(issue_date__range=[start, end]).order_by('issue_date')

        if qs.count() > 1000000:
            base_url = request.scheme + '://' + request.get_host()
            excel_sales_report.delay(month_year, start_date, end_date, airline, is_arc_var, sales_version,
                                     request.user.email, base_url)
            messages.add_message(self.request, messages.WARNING,
                                 'Excel file generation is taking more time than expected. You will receive an email with the link to download the file once it is done.')
            return HttpResponseRedirect('/reports/sales/?' + request.META.get('QUERY_STRING'))

        else:
            response = HttpResponse(content_type='application/vnd.ms-excel')
            airline_obj = Airline.objects.filter(pk=self.request.GET.get('airline', '')).first()
            if airline_obj:
                if month_year:
                    dt_rep = month_year
                else:
                    dt_rep = start_date + " - " + end_date

                file_name = airline_obj.abrev + " " + dt_rep + " Sales Details.xls"
            else:
                file_name = "Sales Details.xls"
            response['Content-Disposition'] = 'inline; filename=' + file_name

            wb = xlwt.Workbook(encoding='utf-8')
            ws = FitSheetWrapper(wb.add_sheet('Sales Details',cell_overwrite_ok=True))

            # Sheet header, first row
            row_num = 2
            month_year = self.request.GET.get('month_year')
         
            font_style = xlwt.XFStyle()
            font_style.font.bold = True


            style_string = "font: bold on; pattern: pattern solid, fore_color yellow; borders: top_color black, bottom_color black, right_color black, left_color black,left thin, right thin, top thin, bottom thin;"
            yellow_style = xlwt.easyxf(style_string)
            stylestr2 = "font: bold on; pattern: pattern solid, fore_color yellow; borders: top_color yellow, bottom_color yellow, right_color yellow, left_color yellow,left thin, right thin, top thin, bottom thin;"
            netsalestyle = xlwt.easyxf(stylestr2)
            default_columns = ["FOP", "Card Code", "Agency", "Date", "Ticket no",
                               "Total",
                               "Fare amount" if not is_arc_var else "Base",
                               "Commission", "Commission %", "Trnc", "PEN", "Tax YQ", "Tax YR"]

            if is_arc_var:
                del default_columns[0]

            peds = qs.values_list('report__report_period__week', flat=True).order_by(
                'report__report_period__week').distinct()

            columns = ["FOP", "Card Code", "Agency", "Date", "Ticket no", "Total",
                       "Fare amount" if not is_arc_var else "Base",
                       "Commission", "Commission %", "Type", "Cancellation penalty"]


            if  self.request.GET.get('sales_version'):
                columns = ["FOP", "Card Code", "Agency","Agency name", "Telephone", "Date", "Ticket no", "Total",
                       "Fare amount" if not is_arc_var else "Base",
                       "Commission", "Commission %", "Type", "Cancellation penalty"]



            if is_arc_var:
                del columns[0]

            if not self.request.GET.get('sales_version'):

                columns = ["FOP", "Card Code", "Agency", "Date", "Ticket no", "Total",
                           "Fare amount" if not is_arc_var else "Base",
                           "Commission", "Commission %", "Type", "Cancellation penalty","Tax YQ", "Tax YR"]
                if is_arc_var:
                    del columns[0]
                taxes = set(qs.values_list("taxes__type", flat=True))

                for tx_type in taxes:

                    if tx_type and tx_type not in ['YQ', 'YR']:
                        columns.append("Tax " + tx_type)
                columns.append("Net Sales")

            for ped in peds:
                print(peds, ">>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>")
                chg = Charges.objects.select_related('transaction').filter(type = 'CP',transaction__report__report_period__week=ped,transaction__report__airline=airline)


                if self.request.GET.get('sales_version'):

                    transaction_type_counts = qs.filter(report__report_period__week=ped).aggregate(
                        tickets=Count('pk', filter=Q(transaction_type='TKTT') | Q(transaction_type='TKT')),
                        refunds=Count('pk', filter=Q(transaction_type='RFND')),
                        exchanges=Count('pk', filter=Q(fop='EX') | Q(transaction_type='EXCH')))



                    ws.write(row_num, 2, "# Tickets = " + str(transaction_type_counts.get('tickets')), font_style)
                    ws.write(row_num, 4, "# Refunds = " + str(transaction_type_counts.get('refunds')), font_style)
                    ws.write(row_num, 6, "# Exchanges = " + str(transaction_type_counts.get('exchanges')), font_style)
                center_normal_bold = xlwt.easyxf(
                "align: wrap yes, vert centre, horiz center;font: name Arial, bold True, height 180")
                ws.write_merge(0,1,0,23,month_year,center_normal_bold)        
                ws.write(row_num, 0, "PED: " + str(ped), font_style)
                row_num = row_num + 1

                for col_num in range(len(columns)):
                    ws.write(row_num, col_num, columns[col_num], yellow_style)

                
                for item in qs.filter(report__report_period__week=ped):

                    amount = 0
                    if not item.transaction_type == "+RTDN":
                        if chg:
                            for c in chg:
                             if item.ticket_no == c.transaction.ticket_no:
                                #  print("cheeeesssss<<<",item.ticket_no,"+++",item.transaction_type,"==",c.transaction.transaction_type,"++",c.amount)
                                 amount =c.amount
                                #  print(">>>>>>>+++",item.ticket_no,"++",c.transaction.ticket_no,"++++",c.amount)

                             pass



                    row_num = row_num + 1
                    tax_start = len(default_columns)
                    col_num = 0
                    if not is_arc_var:
                        ws.write(row_num, col_num, 'CA+CC' if item.ca and item.cc else (item.fop or ''))
                        col_num = col_num + 1

                    ws.write(row_num, col_num, item.card_type or item.card_code)
                    col_num = col_num + 1
                    ws.write(row_num, col_num, item.agency.agency_no or '')
                    col_num = col_num + 1


                    if self.request.GET.get('sales_version'):
                        ws.write(row_num, col_num, item.agency.trade_name or '')
                        col_num = col_num + 1

                        ws.write(row_num, col_num, item.agency.tel or '')
                        col_num = col_num + 1


                    ws.write(row_num, col_num, item.issue_date.strftime("%Y-%m-%d") if item.issue_date else '')
                    col_num = col_num + 1
                    ws.write(row_num, col_num, item.ticket_no or '')
                    col_num = col_num + 1
                    ws.write(row_num, col_num, item.transaction_amount)
                    col_num = col_num + 1
                    ws.write(row_num, col_num, item.fare_amount or 0)
                    col_num = col_num + 1
                    ws.write(row_num, col_num, item.std_comm_amount or 0)
                    col_num = col_num + 1
                    ws.write(row_num, col_num, item.std_comm_rate or 0)
                    col_num = col_num + 1
                    ws.write(row_num, col_num, item.transaction_type or '')
                    col_num = col_num + 1
                    if amount:
                        # cancel_pen=amount
                        ws.write(row_num, col_num, amount or 0)
                        col_num = col_num + 1
                    else:
                      ws.write(row_num, col_num, item.pen or 0)
                      col_num = col_num + 1



                    amount = 0
                    # print("########################################################",item.ticket_no,"+++",amount)
                    # print("pen>>>>>>>>>>>>>",item.transaction_type,"+++=",item.ticket_no,"++++",item.transaction_amount)
                    # ss
                    # ws.write(row_num, col_num, 1 or 0)
                    # col_num = col_num + 1

                    if not self.request.GET.get('sales_version'):

                        ws.write(row_num, col_num, item.yq or 0)
                        col_num = col_num + 1
                        ws.write(row_num, col_num, item.yr or 0)
                        col_num = col_num + 1

                        for tax in taxes:
                            # print(">>>>>>>>d",tax)
                            if tax:
                                tax_amounts = [tx.amount for tx in item.taxes_set.all() if tx.type == tax]
                                if len(tax_amounts) > 0:
                                    tax_amt = sum(tax_amounts)
                                    ws.write(row_num, tax_start, tax_amt)
                                else:
                                    ws.write(row_num, tax_start, 0)

                                tax_start = tax_start + 1
                    col_num = tax_start
                    ws.write(row_num, col_num, item.fare_amount - item.std_comm_amount  or 0,netsalestyle)
                    col_num = col_num + 1

                row_num = row_num + 1
                ws.write(row_num, 0, "Total for report:", yellow_style)
                if not is_arc_var:
                    tot_for_rep = qs.filter(report__report_period__week=ped).aggregate(total=Sum('transaction_amount')).get(
                        'total') #- qs.filter(report__report_period__week=ped).aggregate(
                        #total=Sum('std_comm_amount')).get('total')
                else:
                    tot_for_rep = qs.filter(report__report_period__week=ped).aggregate(total=Sum('transaction_amount')).get(
                        'total') - qs.filter(report__report_period__week=ped).aggregate(
                        total=Sum('std_comm_amount')).get('total')
                if is_arc_var:
                    if self.request.GET.get('sales_version'):
                        ws.write(row_num, 6, format_value(tot_for_rep), yellow_style)
                    else:
                        ws.write(row_num, 4, format_value(tot_for_rep), yellow_style)

                else:
                    if self.request.GET.get('sales_version'):
                        ws.write(row_num, 7, format_value(tot_for_rep), yellow_style)
                    else:
                        ws.write(row_num, 5, format_value(tot_for_rep), yellow_style)
                        
                row_num = row_num + 1
                ws.write(row_num, 0, "")
                row_num = row_num + 1

                # Sheet body, remaining rows

            wb.save(response)
            return response


class GetSalesByReport(PermissionRequiredMixin, View):
    """Filtered Sales By Report download as exel."""

    permission_required = ('report.download_sales_by',)

    def get(self, request, *args, **kwargs):
        organize_by = self.request.GET.get('organize_by', '')
        state = self.request.GET.get('state', '')
        start_date = self.request.GET.get('start_date', '')
        end_date = self.request.GET.get('end_date', '')
        airline = self.request.GET.get('airline', '')
        is_arc_var = is_arc(self.request.session.get('country'))
        qs = Transaction.objects.select_related('agency', 'report').filter(is_sale=True)
        if airline:
            qs = qs.filter(report__airline=airline)

        if state:
            qs = qs.filter(agency__state=state)

        if start_date and end_date:
            start = datetime.datetime.strptime(start_date, '%d %B %Y')
            end = datetime.datetime.strptime(end_date, '%d %B %Y')
            qs = qs.filter(report__report_period__ped__range=[start, end])

        response = HttpResponse(content_type='application/vnd.ms-excel')
        airline_obj = Airline.objects.filter(pk=self.request.GET.get('airline', '')).first()
        if airline_obj:
            dt_rep = start_date + " - " + end_date
            file_name = airline_obj.abrev + " " + dt_rep + "-net_sales_report.xls"
        else:
            file_name = "net_sales_report.xls"
        response['Content-Disposition'] = 'inline; filename=' + file_name

        wb = xlwt.Workbook(encoding='utf-8')
        ws = FitSheetWrapper(wb.add_sheet('Sales By'))
        font_style = xlwt.XFStyle()
        font_style.font.bold = True
        # Sheet header, first row
        row_num = 0

        ws.row(row_num).height = 20 * 20
        ws.write_merge(row_num, 0, 0, 6, airline_obj.name.upper(), bold_center)
        row_num = row_num + 1
        ws.row(row_num).height = 20 * 20
        ws.write_merge(row_num, 1, 0, 6, "SALES BY " + organize_by.upper() + " REPORT", bold_center)
        row_num = row_num + 1
        ws.row(row_num).height = 20 * 20
        ws.write_merge(row_num, 2, 0, 6, "BETWEEN " + start_date.upper() + " AND " + end_date.upper(), bold_center)
        row_num = row_num + 1
        ws.row(row_num).height_mismatch = True

        if organize_by == 'agency':
            qs = qs.values('agency_id').order_by().distinct().annotate(total=Sum('fare_amount'),
                                                                       total_pen=Sum('pen',
                                                                                     filter=Q(pen_type='CANCEL PEN')),
                                                                       agency_trade_name=F('agency__trade_name'),
                                                                       agency_no=F('agency__agency_no'),
                                                                       sales_owner=F('agency__sales_owner__email'),
                                                                       state=F('agency__state__name'),
                                                                       tel=F('agency__tel'),
                                                                       agency_type=F('agency__agency_type__name'))

            grand_sum = qs.aggregate(Sum('total'))
            grand_sum_pen = qs.aggregate(Sum('total_pen'))
            ws.write(row_num, 0, 'Agency ', yellow_background_header)
            ws.write(row_num, 1, 'Sales owner ', yellow_background_header)
            ws.write(row_num, 2, 'Agency trade name ', yellow_background_header)
            ws.write(row_num, 3, 'State ', yellow_background_header)
            ws.write(row_num, 4, 'Tel ', yellow_background_header)
            ws.write(row_num, 5, 'Agency type ', yellow_background_header)
            ws.write(row_num, 6, 'Total ', yellow_background_header)
            row_num = row_num + 1

            for item in qs:
                ws.write(row_num, 0, item.get('agency_no'))
                ws.write(row_num, 1, item.get('sales_owner'))
                ws.write(row_num, 2, item.get('agency_trade_name'))
                ws.write(row_num, 3, item.get('state'))
                ws.write(row_num, 4, item.get('tel'))
                ws.write(row_num, 5, item.get('agency_type'))
                if is_arc_var and item.get('total_pen'):
                    ws.write(row_num, 6, item.get('total') + item.get('total_pen'))
                else:
                    ws.write(row_num, 6, item.get('total'))
                row_num = row_num + 1
            ws.write(row_num, 5, "Total: ", yellow_background_header)
            if is_arc_var and grand_sum_pen.get('total_pen__sum'):
                ws.write(row_num, 6, grand_sum.get('total__sum') + grand_sum_pen.get('total_pen__sum'),
                         yellow_background_header)
            else:
                ws.write(row_num, 6, grand_sum.get('total__sum'), yellow_background_header)

        if organize_by == 'state':
            qs = qs.values('agency__state').order_by().distinct().annotate(total=Sum('fare_amount'),
                                                                           total_pen=Sum('pen',
                                                                                         filter=Q(
                                                                                             pen_type='CANCEL PEN')),
                                                                           sales_owner=F('agency__state__owner__email'),
                                                                           state=F('agency__state__name'))

            grand_sum = qs.aggregate(Sum('total'))
            grand_sum_pen = qs.aggregate(Sum('total_pen'))
            ws.write(row_num, 0, 'State ', yellow_background_header)
            ws.write(row_num, 1, 'State owner ', yellow_background_header)
            ws.write(row_num, 2, 'Total ', yellow_background_header)
            row_num = row_num + 1
            for item in qs:
                ws.write(row_num, 0, item.get('state'))
                ws.write(row_num, 1, item.get('sales_owner'))
                if is_arc_var and item.get('total_pen'):
                    ws.write(row_num, 2, item.get('total') + item.get('total_pen'))
                else:
                    ws.write(row_num, 2, item.get('total'))
                row_num = row_num + 1
            ws.write(row_num, 1, "Total: ", yellow_background_header)

            if is_arc_var and grand_sum_pen.get('total_pen__sum'):
                ws.write(row_num, 2, grand_sum.get('total__sum') + grand_sum_pen.get('total_pen__sum'),
                         yellow_background_header)
            else:
                ws.write(row_num, 2, grand_sum.get('total__sum'), yellow_background_header)

        if organize_by == 'city':
            qs = qs.values('agency__city').order_by().distinct().annotate(total=Sum('fare_amount'),
                                                                          total_pen=Sum('pen',
                                                                                        filter=Q(
                                                                                            pen_type='CANCEL PEN')),
                                                                          state_owner=F('agency__state__owner__email'),
                                                                          city=F('agency__city__name'),
                                                                          state_abrev=F('agency__state__abrev'))

            grand_sum = qs.aggregate(Sum('total'))
            grand_sum_pen = qs.aggregate(Sum('total_pen'))
            ws.write(row_num, 0, 'City ', yellow_background_header)
            ws.write(row_num, 1, 'State Owner ', yellow_background_header)
            ws.write(row_num, 2, 'Total ', yellow_background_header)
            row_num = row_num + 1
            for item in qs:
                ws.write(row_num, 0, (item.get('city') or 'None') + ', ' + (item.get('state_abrev') or 'None'))
                ws.write(row_num, 1, item.get('state_owner'))
                if is_arc_var and item.get('total_pen'):
                    ws.write(row_num, 2, item.get('total') + item.get('total_pen'))
                else:
                    ws.write(row_num, 2, item.get('total'))
                row_num = row_num + 1
            ws.write(row_num, 1, "Total: ", yellow_background_header)

            if is_arc_var and grand_sum_pen.get('total_pen__sum'):
                ws.write(row_num, 2, grand_sum.get('total__sum') + grand_sum_pen.get('total_pen__sum'),
                         yellow_background_header)
            else:
                ws.write(row_num, 2, grand_sum.get('total__sum'), yellow_background_header)

        if organize_by == 'sales owner':
            qs = qs.values('agency__sales_owner').order_by().distinct().annotate(total=Sum('fare_amount'),
                                                                                 total_pen=Sum('pen',
                                                                                               filter=Q(
                                                                                                   pen_type='CANCEL PEN')),
                                                                                 sales_owner=F(
                                                                                     'agency__sales_owner__email'))

            grand_sum = qs.aggregate(Sum('total'))
            grand_sum_pen = qs.aggregate(Sum('total_pen'))
            ws.write(row_num, 0, 'sales owner ', yellow_background_header)
            ws.write(row_num, 1, 'Total ', yellow_background_header)
            row_num = row_num + 1
            for item in qs:
                ws.write(row_num, 0, item.get('sales_owner'))
                if is_arc_var and item.get('total_pen'):
                    ws.write(row_num, 1, item.get('total') + item.get('total_pen'))
                else:
                    ws.write(row_num, 1, item.get('total'))
                row_num = row_num + 1
            ws.write(row_num, 0, "Total: ", yellow_background_header)
            if is_arc_var and grand_sum_pen.get('total_pen__sum'):
                ws.write(row_num, 1, grand_sum.get('total__sum') + grand_sum_pen.get('total_pen__sum'),
                         yellow_background_header)
            else:
                ws.write(row_num, 1, grand_sum.get('total__sum'), yellow_background_header)

        if organize_by == 'agency_type':
            qs = qs.values('agency__agency_type').order_by().distinct().annotate(total=Sum('fare_amount'),
                                                                                 total_pen=Sum('pen',
                                                                                               filter=Q(
                                                                                                   pen_type='CANCEL PEN')),
                                                                                 agency_type=F(
                                                                                     'agency__agency_type__name'))

            grand_sum = qs.aggregate(Sum('total'))
            grand_sum_pen = qs.aggregate(Sum('total_pen'))
            ws.write(row_num, 0, 'agency type ', yellow_background_header)
            ws.write(row_num, 1, 'Total ', yellow_background_header)
            row_num = row_num + 1
            for item in qs:
                ws.write(row_num, 0, item.get('agency_type'))
                if is_arc_var and item.get('total_pen'):
                    ws.write(row_num, 1, item.get('total') + item.get('total_pen'))
                else:
                    ws.write(row_num, 1, item.get('total'))
                row_num = row_num + 1
            ws.write(row_num, 0, "Total: ", yellow_background_header)
            if is_arc_var and grand_sum_pen.get('total_pen__sum'):
                ws.write(row_num, 1, grand_sum.get('total__sum') + grand_sum_pen.get('total_pen__sum'),
                         yellow_background_header)
            else:
                ws.write(row_num, 1, grand_sum.get('total__sum'), yellow_background_header)

        wb.save(response)
        return response


class GetAllSalesReport(PermissionRequiredMixin, View):
    """Filtered All Sales Report download as exel."""

    permission_required = ('report.download_all_sales',)

    def get(self, request, *args, **kwargs):
        sales_type = self.request.GET.get('sales_type', '')
        month_year = self.request.GET.get('month_year', '')
        start_date = self.request.GET.get('start_date', '')
        end_date = self.request.GET.get('end_date', '')

        qs = Transaction.objects.select_related('report__report_period', 'report__airline').exclude(
            transaction_type__startswith='ACM', agency__agency_no='6998051')
        is_arc_var = is_arc(self.request.session.get('country'))

        qs = qs.filter(is_sale=True)

        airline = self.request.GET.get('airline', '')
        if airline:
            qs = qs.filter(report__airline=airline)

        if start_date and end_date:
            start = datetime.datetime.strptime(start_date, '%d %B %Y')
            end = datetime.datetime.strptime(end_date, '%d %B %Y')
            qs = qs.filter(report__report_period__ped__range=[start, end])

        if month_year:

            month = datetime.datetime.strptime(month_year, '%B %Y').month or ''
            year = datetime.datetime.strptime(month_year, '%B %Y').year or ''
            if month and year:
                qs = qs.filter(report__report_period__month=month, report__report_period__year=year)

        airlines = qs.values('report__airline').annotate(name=F('report__airline__name'),
                                                         str_id=Cast(F('report__airline__id'),
                                                                     CharField())).distinct().order_by(
            'report__airline')

        if sales_type == 'Gross':

            annotations = dict()
            for airline in airlines:
                annotation_name = '{}'.format(airline.get('report__airline'))
                annotations[annotation_name] = Sum('transaction_amount',
                                                   filter=Q(report__airline=airline.get('report__airline'))) - Sum(
                    'std_comm_amount', filter=Q(
                        report__airline=airline.get('report__airline')))
            value_list = qs.values('report__report_period__ped').order_by(
                'report__report_period__ped').distinct().annotate(
                **annotations)

            totals = qs.values('report__airline').annotate(
                total=Sum('transaction_amount') - Sum('std_comm_amount')).order_by('report__airline')

        else:

            sub_tax_yq_yr = Charges.objects.select_related('transaction').filter(type__in=['YQ', 'YR'],
                                                                                 transaction__report__report_period__ped=OuterRef(
                                                                                     'report__report_period__ped'),
                                                                                 transaction__report__airline=airline,
                                                                                 transaction__is_sale=True).values(
                'transaction__report__report_period__ped').annotate(
                dcount=Count('transaction__report__report_period__ped'),
                total_tax_yq_yr=Sum('amount', output_field=FloatField())).values('total_tax_yq_yr')
            sub_tax = Taxes.objects.select_related('transaction').filter(
                transaction__report__report_period__ped=OuterRef('report__report_period__ped'),
                transaction__report__airline=airline, transaction__is_sale=True).values(
                'transaction__report__report_period__ped').annotate(
                dcount=Count('transaction__report__report_period__ped'),
                total_tax=Sum('amount', output_field=FloatField())).values('total_tax')

            if not is_arc_var:

                vals = qs.values('report__airline', 'report__report_period__ped').order_by('report__airline',
                                                                                           'report__report_period__ped').annotate(
                    airline_name=F('report__airline__name'),
                    net_sale=Sum('fare_amount') - Sum('std_comm_amount') + Coalesce(Sum('pen'), V(0)),
                    yqyr=Coalesce(Subquery(sub_tax_yq_yr, output_field=FloatField()), V(0)),
                    taxes=Coalesce(Subquery(sub_tax, output_field=FloatField()), V(0)),
                    total=Sum('transaction_amount') - Sum('std_comm_amount'))
                totals = vals.aggregate(yqyr_total=Sum('yqyr'), net_total=Sum('net_sale'),
                                        tax_total=Sum('taxes'), gross_total=Sum('total'))
            else:
                sub_acm = Transaction.objects.select_related('report', 'report__airline',
                                                             'report__report_period').filter(
                    report__airline=airline, transaction_type__startswith='ACM', agency__agency_no='31768026',
                    report__report_period__ped=OuterRef('report__report_period__ped')).values(
                    'report__report_period__ped').annotate(
                    dcount=Count('report__report_period__ped'),
                    total_ap_acm=Sum('transaction_amount', output_field=FloatField())).values('total_ap_acm')
                vals = qs.values('report__airline', 'report__report_period__ped').order_by('report__airline',
                                                                                           'report__report_period__ped').annotate(
                    airline_name=F('report__airline__name'),
                    net_sale=Coalesce(Sum('fare_amount'), V(0)) + Coalesce(Sum('pen'), V(0)),
                    yqyr=Coalesce(Subquery(sub_tax_yq_yr, output_field=FloatField()), V(0)),
                    taxes=Coalesce(Subquery(sub_tax, output_field=FloatField()), V(0)),
                    total=Sum('transaction_amount') - Sum('std_comm_amount') + Coalesce(
                        Subquery(sub_acm, output_field=FloatField()), V(0)))
                totals = vals.aggregate(yqyr_total=Sum('yqyr'), net_total=Sum('net_sale'),
                                        tax_total=Sum('taxes'), gross_total=Sum('total'))

        response = HttpResponse(content_type='application/vnd.ms-excel')

        wb = xlwt.Workbook(encoding='utf-8')
        ws = FitSheetWrapper(wb.add_sheet('All Sales for month'))
        ws.col(0).width = 250 * PT
        font_style = xlwt.XFStyle()
        font_style.font.bold = True
        # Sheet header, first row
        row_num = 0
        airlineName = ''
        if month_year:
            file_name = month_year + "-all-sales.xls"
            if self.request.GET.get('airline', ''):
                ws.row(row_num).height = 20 * 20
                for i, airline in enumerate(airlines):
                    airlineName = airline.get('name')

                ws.write_merge(row_num, 0, 0, 7, airlineName.upper(), bold_center)
                row_num = row_num + 1
                ws.row(row_num).height = 20 * 20
                ws.write_merge(row_num, 1, 0, 7, sales_type.upper() + " SALES", bold_center)
                row_num = row_num + 1
                ws.row(row_num).height = 20 * 20
                ws.write_merge(row_num, 2, 0, 7, month_year.upper(), bold_center)
                row_num = row_num + 1
                ws.row(row_num).height_mismatch = True
            else:
                ws.row(row_num).height = 20 * 20
                ws.write_merge(row_num, 0, 0, 7, sales_type.upper() + " SALES", bold_center)
                row_num = row_num + 1
                ws.row(row_num).height = 20 * 20
                ws.write_merge(row_num, 1, 0, 7, month_year.upper(), bold_center)
                row_num = row_num + 1
                ws.row(row_num).height_mismatch = True
        else:
            file_name = start_date + ' - ' + end_date + "-all-sales.xls"

            if self.request.GET.get('airline', ''):
                ws.row(row_num).height = 20 * 20
                for i, airline in enumerate(airlines):
                    airlineName = airline.get('name')

                ws.write_merge(row_num, 0, 0, 7, airlineName.upper(), bold_center)
                row_num = row_num + 1
                ws.row(row_num).height = 20 * 20
                ws.write_merge(row_num, 1, 0, 7, sales_type.upper() + " SALES", bold_center)
                row_num = row_num + 1
                ws.row(row_num).height = 20 * 20
                ws.write_merge(row_num, 2, 0, 7, 'FROM ' + start_date.upper() + ' TO ' + end_date.upper(), bold_center)
                row_num = row_num + 1
                ws.row(row_num).height_mismatch = True
            else:
                ws.row(row_num).height = 20 * 20
                ws.write_merge(row_num, 0, 0, 7, sales_type.upper() + " SALES", bold_center)
                row_num = row_num + 1
                ws.row(row_num).height = 20 * 20
                ws.write_merge(row_num, 1, 0, 7, 'FROM ' + start_date.upper() + ' TO ' + end_date.upper(), bold_center)
                row_num = row_num + 1
                ws.row(row_num).height_mismatch = True

        response['Content-Disposition'] = 'inline; filename=' + file_name

        if sales_type == 'Gross':

            ws.write(row_num, 0, 'PED ', yellow_background_header)
            for i, airline in enumerate(airlines):
                ws.write(row_num, i + 1, airline.get('name'), yellow_background_header)
            row_num = row_num + 1

            for value in value_list:
                ws.write(row_num, 0, value.get('report__report_period__ped').strftime("%Y-%m-%d"))
                for i, airline in enumerate(airlines):
                    ws.write(row_num, i + 1, value.get(airline.get('str_id')) or 0.0)
                row_num = row_num + 1

            ws.write(row_num, 0, 'Total : ', yellow_background_header)
            for i, total in enumerate(totals):
                ws.write(row_num, i + 1, total.get('total'), yellow_background_header)
        else:
            ws.write(row_num, 0, 'Airline ', yellow_background_header)
            ws.write(row_num, 1, 'PED ', yellow_background_header)
            ws.write(row_num, 2, 'Net Sales ', yellow_background_header)
            ws.write(row_num, 3, 'YQ+YR ', yellow_background_header)
            ws.write(row_num, 4, 'Other Tax ', yellow_background_header)
            ws.write(row_num, 5, 'Gross Sales ', yellow_background_header)

            row_num = row_num + 1

            for val in vals:
                ws.write(row_num, 0, val.get('airline_name'))
                ws.write(row_num, 1, val.get('report__report_period__ped').strftime("%Y-%m-%d"))
                ws.write(row_num, 2, val.get('net_sale'))
                ws.write(row_num, 3, val.get('yqyr'))
                ws.write(row_num, 4, val.get('taxes'))
                ws.write(row_num, 5, val.get('total'))
                row_num = row_num + 1

            ws.write(row_num, 0, 'Total: ', yellow_background_header)
            ws.write(row_num, 2, totals.get('net_total'), yellow_background_header)
            ws.write(row_num, 3, totals.get('yqyr_total'), yellow_background_header)
            ws.write(row_num, 4, totals.get('tax_total'), yellow_background_header)
            ws.write(row_num, 5, totals.get('gross_total'), yellow_background_header)
        wb.save(response)
        return response


class GetYearToYearSalesReport(PermissionRequiredMixin, View):
    """Filtered year to year Sales Report download as exel."""

    permission_required = ('report.download_year_to_years',)

    def get(self, request, *args, **kwargs):
        month_year = self.request.GET.get('month_year', '')
        sales_type = self.request.GET.get('sales_type', '')
        qs = Transaction.objects.select_related('report__report_period', 'report__airline').filter(is_sale=True)
        airlines = Airline.objects.filter(country=self.request.session.get('country'))
        is_arc_var = is_arc(self.request.session.get('country'))

        if is_arc_var:
            qs_disp = Disbursement.objects.select_related('report_period', 'airline')
        else:
            qs_acm = Transaction.objects.select_related('report__report_period', 'report__airline', 'agency').filter(
                transaction_type__startswith='ACM', agency__agency_no='6998051', is_sale=True)

        if month_year:

            month = datetime.datetime.strptime(month_year, '%B %Y').month or ''
            year = datetime.datetime.strptime(month_year, '%B %Y').year or ''
            month_year2 = datetime.datetime(year, month, 1).strftime("%b %Y")
            month_year1 = (datetime.datetime(year, month, 1) - rd(years=1)).strftime("%b %Y")
            values = []

            if month and year:
                qs1 = qs.filter(report__report_period__month=month, report__report_period__year=year - 1)
                qs2 = qs.filter(report__report_period__month=month, report__report_period__year=year)
                if is_arc_var:
                    qs1_disp = qs_disp.filter(report_period__month=month, report_period__year=year - 1)
                    qs2_disp = qs_disp.filter(report_period__month=month, report_period__year=year)
                else:
                    qs1_acm = qs_acm.filter(report__report_period__month=month, report__report_period__year=year - 1)
                    qs2_acm = qs_acm.filter(report__report_period__month=month, report__report_period__year=year)

                airlines = airlines.filter(
                    id__in=qs.filter(report__report_period__year__in=[year - 1, year],
                                     report__report_period__month=month).values_list(
                        'report__airline', flat=True))
                for airline in airlines:
                    sales_data = {}
                    sales_data['airline'] = airline

                    if sales_type == 'Net':
                        if is_arc_var:
                            acm_y1 = Transaction.objects.select_related('report', 'report__airline',
                                                                        'report__report_period').filter(
                                report__airline=airline, transaction_type__startswith='ACM',
                                agency__agency_no='31768026', report__report_period__month=month,
                                report__report_period__year=year - 1).aggregate(
                                total_ap_acm=Coalesce(Sum('transaction_amount'), V(0))).get('total_ap_acm')
                            acm_y2 = Transaction.objects.select_related('report', 'report__airline',
                                                                        'report__report_period').filter(
                                report__airline=airline, transaction_type__startswith='ACM',
                                agency__agency_no='31768026', report__report_period__month=month,
                                report__report_period__year=year).aggregate(
                                total_ap_acm=Coalesce(Sum('transaction_amount'), V(0))).get(
                                'total_ap_acm')

                            sales_data['amount1'] = qs1.filter(report__airline=airline).aggregate(
                                amount1=Coalesce(Sum('fare_amount'), V(0)) + Coalesce(Sum('pen'), V(0))).get(
                                'amount1') - acm_y1 or 0
                            sales_data['cash1'] = qs1_disp.filter(airline=airline).aggregate(
                                cash1=Sum('bank7')).get('cash1') or 0
                            sales_data['amount2'] = qs2.filter(report__airline=airline).aggregate(
                                amount2=Coalesce(Sum('fare_amount'), V(0)) + Coalesce(Sum('pen'), V(0))).get(
                                'amount2') - acm_y2 or 0
                            sales_data['cash2'] = qs2_disp.filter(airline=airline).aggregate(
                                cash1=Sum('bank7')).get('cash1') or 0
                        else:
                            qs1_acm_nw = qs1_acm.filter(report__airline=airline).aggregate(
                                total=Sum('fare_amount', output_field=FloatField()),
                                total_comm=Sum('std_comm_amount', output_field=FloatField()),
                                total_cp=Sum('pen', output_field=FloatField()))

                            qs2_acm_nw = qs2_acm.filter(report__airline=airline).aggregate(
                                total=Sum('fare_amount', output_field=FloatField()),
                                total_comm=Sum('std_comm_amount', output_field=FloatField()),
                                total_cp=Sum('pen', output_field=FloatField()))
                            sales_data['amount1'] = qs1.filter(report__airline=airline).aggregate(
                                amount1=(Sum('fare_amount') - Coalesce(qs1_acm_nw.get('total'), V(0))) - (
                                    Sum('std_comm_amount', output_field=FloatField()) - Coalesce(
                                    qs1_acm_nw.get('total_comm'), V(0))) + (
                                            Sum('pen') - Coalesce(qs1_acm_nw.get('total_cp'), V(0)))).get(
                                'amount1') or 0
                            sales_data['cash1'] = qs1.filter(report__airline=airline).aggregate(
                                cash1=Sum('ca') - (Sum('std_comm_amount', output_field=FloatField()) - Coalesce(
                                    qs1_acm_nw.get('total_comm'), V(0)))).get(
                                'cash1') or 0
                            sales_data['amount2'] = qs2.filter(report__airline=airline).aggregate(
                                amount2=(Sum('fare_amount') - Coalesce(qs2_acm_nw.get('total'), V(0))) - (
                                    Sum('std_comm_amount', output_field=FloatField()) - Coalesce(
                                    qs2_acm_nw.get('total_comm'), V(0))) + (
                                            Sum('pen') - Coalesce(qs2_acm_nw.get('total_cp'), V(0)))).get(
                                'amount2') or 0
                            sales_data['cash2'] = qs2.filter(report__airline=airline).aggregate(
                                cash2=Sum('ca') - (Sum('std_comm_amount', output_field=FloatField()) - Coalesce(
                                    qs2_acm_nw.get('total_comm'), V(0)))).get(
                                'cash2') or 0
                    else:
                        if is_arc_var:
                            sales_data['amount1'] = qs1.filter(report__airline=airline).aggregate(
                                amount1=Sum('transaction_amount', output_field=FloatField()) - Sum('std_comm_amount',
                                                                                                   output_field=FloatField())).get(
                                'amount1') or 0
                            sales_data['amount2'] = qs2.filter(report__airline=airline).aggregate(
                                amount2=Sum('transaction_amount', output_field=FloatField()) - Sum('std_comm_amount',
                                                                                                   output_field=FloatField())).get(
                                'amount2') or 0
                            sales_data['cash1'] = qs1_disp.filter(airline=airline).aggregate(
                                cash1=Sum('bank7')).get('cash1') or 0
                            sales_data['cash2'] = qs2_disp.filter(airline=airline).aggregate(
                                cash1=Sum('bank7')).get('cash1') or 0
                        else:
                            qs1_acm_nw = qs1_acm.filter(report__airline=airline).aggregate(
                                total=Sum('transaction_amount', output_field=FloatField()),
                                total_comm=Sum('std_comm_amount', output_field=FloatField()))

                            qs2_acm_nw = qs2_acm.filter(report__airline=airline).aggregate(
                                total=Sum('transaction_amount', output_field=FloatField()),
                                total_comm=Sum('std_comm_amount', output_field=FloatField()))

                            sales_data['amount1'] = qs1.filter(report__airline=airline).aggregate(
                                amount1=(Sum('transaction_amount', output_field=FloatField()) - Coalesce(
                                    qs1_acm_nw.get('total'), V(0))) - (
                                            Sum('std_comm_amount', output_field=FloatField()) - Coalesce(
                                            qs1_acm_nw.get('total_comm'), V(0)))).get(
                                'amount1') or 0
                            sales_data['amount2'] = qs2.filter(report__airline=airline).aggregate(
                                amount2=(Sum('transaction_amount', output_field=FloatField()) - Coalesce(
                                    qs2_acm_nw.get('total'), V(0))) - (
                                            Sum('std_comm_amount', output_field=FloatField()) - Coalesce(
                                            qs2_acm_nw.get('total_comm'), V(0)))).get(
                                'amount2') or 0
                            sales_data['cash1'] = qs1.filter(report__airline=airline).aggregate(
                                cash1=Sum('ca') - (Sum('std_comm_amount', output_field=FloatField()) - Coalesce(
                                    qs1_acm_nw.get('total_comm'), V(0)))).get(
                                'cash1') or 0
                            sales_data['cash2'] = qs2.filter(report__airline=airline).aggregate(
                                cash2=Sum('ca') - (Sum('std_comm_amount', output_field=FloatField()) - Coalesce(
                                    qs2_acm_nw.get('total_comm'), V(0)))).get(
                                'cash2') or 0

                    values.append(sales_data)
        sales_type = self.request.GET.get('sales_type', '')

        response = HttpResponse(content_type='application/vnd.ms-excel')

        wb = xlwt.Workbook(encoding='utf-8')
        ws = FitSheetWrapper(wb.add_sheet('All Sales for month'))
        ws.col(0).width = 250 * PT
        font_style = xlwt.XFStyle()
        font_style.font.bold = True
        # Sheet header, first row
        row_num = 0
        if month_year:
            file_name = month_year + "-year-to-year-sales.xls"

            ws.row(row_num).height = 20 * 20
            ws.write_merge(row_num, 0, 0, 7, sales_type.upper() + " SALES FIGURES", bold_center)
            row_num = row_num + 1
        else:
            file_name = "year-to-year-sales.xls"
            ws.row(row_num).height = 20 * 20
            ws.write_merge(row_num, 0, 0, 7, sales_type.upper() + " YEAR TO YEAR SALES", bold_center)
            row_num = row_num + 1

        ws.row(row_num).height = 20 * 20
        ws.write_merge(row_num, 1, 0, 7, month_year.upper(), bold_center)
        row_num = row_num + 1
        ws.row(row_num).height_mismatch = True

        response['Content-Disposition'] = 'inline; filename=' + file_name

        ws.write(row_num, 0, 'Airline ', yellow_background_header)
        ws.write(row_num, 1, month_year1, yellow_background_header)
        ws.write(row_num, 2, 'Cash', yellow_background_header)
        ws.write(row_num, 3, month_year2, yellow_background_header)
        ws.write(row_num, 4, 'Cash', yellow_background_header)

        row_num = row_num + 1

        for val in values:
            ws.write(row_num, 0, val.get('airline').name)
            ws.write(row_num, 1, val.get('amount1'))
            ws.write(row_num, 2, val.get('cash1'))
            ws.write(row_num, 3, val.get('amount2'))
            ws.write(row_num, 4, val.get('cash2'))
            row_num = row_num + 1

        wb.save(response)
        return response


@task(time_limit=999999, soft_time_limit=999999)
def excel_commission_report(start_date, end_date, month_year, airline, to_email, base_url):
    if airline:
        qs = Transaction.objects.select_related('agency__city', 'agency__state', 'report__airline').exclude(
            std_comm_rate__isnull=True).order_by('report__report_period__ped', 'report__airline').filter(
            report__airline=airline, is_sale=True)
        if month_year:
            month = datetime.datetime.strptime(month_year, '%B %Y').month or ''
            year = datetime.datetime.strptime(month_year, '%B %Y').year or ''
            if month and year:
                qs = qs.filter(report__report_period__month=month, report__report_period__year=year)

        if start_date and end_date:
            start = datetime.datetime.strptime(start_date, '%d %B %Y')
            end = datetime.datetime.strptime(end_date, '%d %B %Y')
            qs = qs.filter(report__report_period__ped__range=[start, end])

        airline_obj = Airline.objects.filter(pk=airline).first()
        airline_name = airline_obj.name
        wb = xlwt.Workbook(encoding='utf-8')
        ws = FitSheetWrapper(wb.add_sheet('Commission Report'))
        row_num = 0
        if airline_obj:
            if month_year:
                dt_rep = month_year

                ws.row(row_num).height = 20 * 20
                ws.write_merge(row_num, 0, 0, 7, airline_obj.name.upper(), bold_center)
                row_num = row_num + 1
                ws.row(row_num).height = 20 * 20
                ws.write_merge(row_num, 1, 0, 7, "COMMISSION REPORT", bold_center)
                row_num = row_num + 1
                ws.row(row_num).height = 20 * 20
                ws.write_merge(row_num, 2, 0, 7, month_year.upper(), bold_center)
                row_num = row_num + 1
                ws.row(row_num).height_mismatch = True
            else:
                dt_rep = start_date + " - " + end_date

                ws.row(row_num).height = 20 * 20
                ws.write_merge(row_num, 0, 0, 7, "COMMISSION REPORT", bold_center)
                row_num = row_num + 1
                ws.row(row_num).height = 20 * 20
                ws.write_merge(row_num, 1, 0, 7, month_year.upper(), bold_center)
                row_num = row_num + 1
                ws.row(row_num).height_mismatch = True

            file_name = airline_obj.abrev + " " + dt_rep + " commission_report.xls"
        else:
            file_name = "commission_report.xls"

        peds = qs.values_list('report__report_period__ped', flat=True).order_by(
            'report__report_period__ped').distinct()

        for ped in peds:
            ws.write(row_num, 0, "PED: " + str(ped), yellow_background_header)
            row_num = row_num + 2
            columns = ['Agency', 'Agency name', 'Agency address', 'Agency city', 'Agency state',
                       'Agency phone number',
                       'Ticket number', 'Commission taken (%)', 'Max sales commission (%)']
            for col_num in range(len(columns)):
                ws.write(row_num, col_num, columns[col_num], yellow_background_header)

            trns_list = []
            commissions_history = CommissionHistory.objects.filter(airline=airline, type='M').values('from_date',
                                                                                                     'to_date',
                                                                                                     'rate').order_by(
                '-from_date')
            for transaction in qs.filter(report__report_period__ped=ped).order_by('agency'):
                histories = [history for history in commissions_history if
                             history.get('from_date') <= transaction.report.report_period.ped]
                rate = histories[0].get('rate') or 0.0 if histories else 0.0
                if transaction.std_comm_rate < rate:
                    trn = dict()
                    trn['agency_no'] = transaction.agency.agency_no
                    trn['agency_id'] = transaction.agency.pk
                    trn['agency_name'] = transaction.agency.trade_name
                    trn['agency_address'] = transaction.agency.address1
                    trn['agency_city'] = transaction.agency.city.name if transaction.agency.city else ''
                    trn['agency_state'] = transaction.agency.state.name if transaction.agency.state else ''
                    trn['agency_tel'] = transaction.agency.tel
                    trn['ticket_no'] = transaction.ticket_no
                    trn['ped'] = transaction.report.report_period.ped
                    trn['std_comm_rate'] = transaction.std_comm_rate
                    trn['max_comm_rate'] = rate
                    trns_list.append(trn)

            for item in trns_list:
                row_num = row_num + 1
                ws.write(row_num, 0, item.get('agency_no'))
                ws.write(row_num, 1, item.get('agency_name'))
                ws.write(row_num, 2, item.get('agency_address'))
                ws.write(row_num, 3, item.get('agency_city'))
                ws.write(row_num, 4, item.get('agency_state'))
                ws.write(row_num, 5, item.get('agency_tel'))
                ws.write(row_num, 6, item.get('ticket_no'))
                ws.write(row_num, 7, item.get('std_comm_rate'))
                ws.write(row_num, 8, item.get('max_comm_rate'))
            row_num = row_num + 2

            # Sheet body, remaining rows

        path = os.path.join(settings.MEDIA_ROOT, 'excelreports', file_name)
        wb.save(path)

        excel = ExcelReportDownload()
        excel.report_type = 2
        excel.file.name = 'excelreports/' + file_name
        excel.save()
        context = {'user': to_email,
                   'link': base_url + excel.file.url,
                   'report_name': 'Commission Report',
                   'airline_name': airline_name,
                   'date': month_year
                   }
        send_mail('Commission Report ' + month_year + ': ' + airline_name, "email/excel-report-email.html", context,
                  [to_email], from_email='assda@assda.com')
    return True


class GetCommissionReport(PermissionRequiredMixin, View):
    """Filtered Commission Report download as exel."""

    permission_required = ('report.download_commission',)

    def get(self, request, *args, **kwargs):
        month_year = self.request.GET.get('month_year', '')
        start_date = self.request.GET.get('start_date', '')
        end_date = self.request.GET.get('end_date', '')
        airline = self.request.GET.get('airline', '')

        if airline:
            qs = Transaction.objects.select_related('agency__city', 'agency__state', 'report__airline').exclude(
                std_comm_rate__isnull=True).order_by('report__report_period__ped', 'report__airline').filter(
                report__airline=airline, is_sale=True)
            if month_year:
                month = datetime.datetime.strptime(month_year, '%B %Y').month or ''
                year = datetime.datetime.strptime(month_year, '%B %Y').year or ''
                if month and year:
                    qs = qs.filter(report__report_period__month=month, report__report_period__year=year)

            if start_date and end_date:
                start = datetime.datetime.strptime(start_date, '%d %B %Y')
                end = datetime.datetime.strptime(end_date, '%d %B %Y')
                qs = qs.filter(report__report_period__ped__range=[start, end])

            if qs.count() > 10000:
                base_url = request.scheme + '://' + request.get_host()
                excel_commission_report.delay(start_date, end_date, month_year, airline, request.user.email, base_url)
                messages.add_message(self.request, messages.WARNING,
                                     'Excel file generation is taking more time than expected. You will receive an email with the link to download the file once it is done.')
                return HttpResponseRedirect('/reports/commission?' + request.META.get('QUERY_STRING'))

            response = HttpResponse(content_type='application/vnd.ms-excel')
            airline_obj = Airline.objects.filter(pk=self.request.GET.get('airline', '')).first()
            wb = xlwt.Workbook(encoding='utf-8')
            ws = FitSheetWrapper(wb.add_sheet('Commission Report'))
            row_num = 0
            if airline_obj:
                if month_year:
                    dt_rep = month_year

                    ws.row(row_num).height = 20 * 20
                    ws.write_merge(row_num, 0, 0, 7, airline_obj.name.upper(), bold_center)
                    row_num = row_num + 1
                    ws.row(row_num).height = 20 * 20
                    ws.write_merge(row_num, 1, 0, 7, "COMMISSION REPORT", bold_center)
                    row_num = row_num + 1
                    ws.row(row_num).height = 20 * 20
                    ws.write_merge(row_num, 2, 0, 7, month_year.upper(), bold_center)
                    row_num = row_num + 1
                    ws.row(row_num).height_mismatch = True
                else:
                    dt_rep = start_date + " - " + end_date

                    ws.row(row_num).height = 20 * 20
                    ws.write_merge(row_num, 0, 0, 7, "COMMISSION REPORT", bold_center)
                    row_num = row_num + 1
                    ws.row(row_num).height = 20 * 20
                    ws.write_merge(row_num, 1, 0, 7, month_year.upper(), bold_center)
                    row_num = row_num + 1
                    ws.row(row_num).height_mismatch = True

                file_name = airline_obj.abrev + " " + dt_rep + " commission_report.xls"
            else:
                file_name = "commission_report.xls"
            response['Content-Disposition'] = 'inline; filename=' + file_name

            peds = qs.values_list('report__report_period__ped', flat=True).order_by(
                'report__report_period__ped').distinct()

            for ped in peds:
                ws.write(row_num, 0, "PED: " + str(ped), yellow_background_header)
                row_num = row_num + 2
                columns = ['Agency', 'Agency name', 'Agency address', 'Agency city', 'Agency state',
                           'Agency phone number',
                           'Ticket number', 'Commission taken (%)', 'Max sales commission (%)']
                for col_num in range(len(columns)):
                    ws.write(row_num, col_num, columns[col_num], yellow_background_header)

                trns_list = []
                commissions_history = CommissionHistory.objects.filter(airline=airline, type='M').values('from_date',
                                                                                                         'to_date',
                                                                                                         'rate').order_by(
                    '-from_date')
                for transaction in qs.filter(report__report_period__ped=ped).order_by('agency'):
                    histories = [history for history in commissions_history if
                                 history.get('from_date') <= transaction.report.report_period.ped]
                    rate = histories[0].get('rate') or 0.0 if histories else 0.0
                    if transaction.std_comm_rate < rate:
                        trn = dict()
                        trn['agency_no'] = transaction.agency.agency_no
                        trn['agency_id'] = transaction.agency.pk
                        trn['agency_name'] = transaction.agency.trade_name
                        trn['agency_address'] = transaction.agency.address1
                        trn['agency_city'] = transaction.agency.city.name if transaction.agency.city else ''
                        trn['agency_state'] = transaction.agency.state.name if transaction.agency.state else ''
                        trn['agency_tel'] = transaction.agency.tel
                        trn['ticket_no'] = transaction.ticket_no
                        trn['ped'] = transaction.report.report_period.ped
                        trn['std_comm_rate'] = transaction.std_comm_rate
                        trn['max_comm_rate'] = rate
                        trns_list.append(trn)

                for item in trns_list:
                    row_num = row_num + 1
                    ws.write(row_num, 0, item.get('agency_no'))
                    ws.write(row_num, 1, item.get('agency_name'))
                    ws.write(row_num, 2, item.get('agency_address'))
                    ws.write(row_num, 3, item.get('agency_city'))
                    ws.write(row_num, 4, item.get('agency_state'))
                    ws.write(row_num, 5, item.get('agency_tel'))
                    ws.write(row_num, 6, item.get('ticket_no'))
                    ws.write(row_num, 7, item.get('std_comm_rate'))
                    ws.write(row_num, 8, item.get('max_comm_rate'))
                row_num = row_num + 2

                # Sheet body, remaining rows

            wb.save(response)
            return response
        return HttpResponse('something went wrong')


class GetTopAgentReport(PermissionRequiredMixin, View):
    """Filtered Top Agent Report download as exel."""

    permission_required = ('report.download_top_agency',)

    def get(self, request, *args, **kwargs):
        start_date = self.request.GET.get('start_date', '')
        end_date = self.request.GET.get('end_date', '')
        airline = self.request.GET.get('airline', '')
        limit = int(self.request.GET.get('limit', 1))
        selected_city = self.request.GET.get('city', '')
        selected_state = self.request.GET.get('state', '')

        qs = Transaction.objects.select_related('agency__city', 'agency__state', 'report__airline').exclude(
            transaction_type__startswith='ACM', agency__agency_no='6998051')

        qs = qs.filter(is_sale=True)

        if airline:
            qs = qs.filter(report__airline=airline)

        if selected_state:
            qs = qs.filter(agency__state=selected_state)
        if selected_city:
            qs = qs.filter(agency__city=selected_city)

        if start_date and end_date:
            start = datetime.datetime.strptime(start_date, '%d %B %Y')
            end = datetime.datetime.strptime(end_date, '%d %B %Y')
            qs = qs.filter(report__report_period__ped__range=[start, end])

        qs = qs.values('agency').annotate(
            fare_sum=Coalesce(Sum('fare_amount'), V(0)) - Coalesce(Sum('std_comm_amount'), V(0)) + Coalesce(Sum('pen'),
                                                                                                            V(0)),
            agency_name=F('agency__trade_name'),
            agency_city=F('agency__city__name'), agency_state=F('agency__state__name'),
            tel=F('agency__tel'), agency_email=F('agency__email'),
            agency_no=F('agency__agency_no'),
            sales_owner=F('agency__sales_owner__email')).order_by('-fare_sum')[:limit]

        response = HttpResponse(content_type='application/vnd.ms-excel')
        airline_obj = Airline.objects.filter(pk=self.request.GET.get('airline', '')).first()
        wb = xlwt.Workbook(encoding='utf-8')
        ws = FitSheetWrapper(wb.add_sheet('Top Agency Report'))
        row_num = 0
        if airline_obj:
            dt_rep = start_date + " - " + end_date


            ws.row(row_num).height = 20 * 20
            ws.write_merge(row_num, 0, 0, 7, airline_obj.name.upper(), bold_center)
            row_num = row_num + 1
            ws.row(row_num).height = 20 * 20
            ws.write_merge(row_num, 1, 0, 7, "TOP AGENCY REPORT", bold_center)
            row_num = row_num + 1
            ws.row(row_num).height = 20 * 20
            ws.write_merge(row_num, 2, 0, 7, "FROM " + start_date.upper() + ' TO ' + end_date.upper(), bold_center)
            row_num = row_num + 1
            ws.row(row_num).height_mismatch = True

            file_name = airline_obj.abrev + " " + dt_rep + " top_agent_report.xls"
        else:
            file_name = "top_agent_report.xls"
        response['Content-Disposition'] = 'inline; filename=' + file_name

        ws.write(row_num, 0, "Sales Owner", yellow_background_header)
        ws.write(row_num, 1, "Agency number ", yellow_background_header)
        ws.write(row_num, 2, "Agency name", yellow_background_header)
        ws.write(row_num, 3, "City", yellow_background_header)
        ws.write(row_num, 4, "State", yellow_background_header)
        ws.write(row_num, 5, "Tel", yellow_background_header)
        ws.write(row_num, 6, "Email", yellow_background_header)
        ws.write(row_num, 7, "Sales", yellow_background_header)

        row_num = row_num + 1

        for agent in qs:
            ws.write(row_num, 0, agent.get('sales_owner') or '')
            ws.write(row_num, 1, agent.get('agency_no') or '')
            ws.write(row_num, 2, agent.get('agency_name') or '')
            ws.write(row_num, 3, agent.get('agency_city') or '')
            ws.write(row_num, 4, agent.get('agency_state') or '')
            ws.write(row_num, 5, agent.get('tel') or '')
            ws.write(row_num, 6, agent.get('agency_email') or '')
            ws.write(row_num, 7, agent.get('fare_sum') or 0)
            row_num = row_num + 1

            # Sheet body, remaining rows

        wb.save(response)
        return response


class ReProcessReports(PermissionRequiredMixin, View):
    """Reprocess report files."""

    permission_required = ('report.view_upload_reports',)
    template_name = 're-process.html'

    def get(self, request, *args, **kwargs):
        context = dict()
        context['airlines'] = Airline.objects.filter(country=self.request.session.get('country'))
        context['pendings'] = ReprocessFile.objects.filter(is_done=False)
        return render(request, self.template_name, context)

    def post(self, request, *args, **kwargs):
        start_date = self.request.POST.get('start_date', '')
        end_date = self.request.POST.get('end_date', '')
        airline = self.request.POST.get('airline', '')
        context = dict()
        context['airlines'] = Airline.objects.filter(country=self.request.session.get('country'))
        context['pendings'] = ReprocessFile.objects.filter(is_done=False)
        if ReprocessFile.objects.filter(is_done=False):
            messages.add_message(self.request, messages.ERROR, 'Already a process running.')
            return render(request, self.template_name, context)
        if start_date and end_date:
            task = ReprocessFile()
            if airline:
                task.airline = Airline.objects.get(pk=airline)
            task.start_date = datetime.datetime.strptime(start_date, '%d %B %Y')
            task.end_date = datetime.datetime.strptime(end_date, '%d %B %Y')
            task.save()

            re_process.delay(airline, start_date, end_date, task.id)

        return redirect('re_process')


class CheckTasks(PermissionRequiredMixin, View):
    permission_required = ('report.view_upload_reports',)

    def get(self, request):
        if ReprocessFile.objects.filter(is_done=False).exists():
            return JsonResponse({'has_pending': 'true'})
        return JsonResponse({'has_pending': 'false'})


class CalendarUpload(PermissionRequiredMixin, View):
    """Calendar upload"""

    permission_required = ('report.view_upload_calendar',)
    template_name = 'calendar-uploads.html'
    success_url = '/reports/upload/'

    def get(self, request, *args, **kwargs):
        context = dict()
        return render(request, self.template_name, context)

    def post(self, request):
        file = request.FILES['file']
        file_name, extention = file.name.split('.')

        if extention.lower() in ['xlsx']:
            resp = process_excelfile(file, request)
        else:
            resp = "Invalid calendar format, please upload xlsx file."

        if resp == 'success':
            messages.add_message(self.request, messages.SUCCESS, 'Calendar file uploaded successfully')
        else:
            messages.add_message(self.request, messages.ERROR, resp)

        return render(request, self.template_name)


class CalendarList(PermissionRequiredMixin, ListView):
    """Monthly YOY report ."""
    months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October',
              'November', 'December']
    permission_required = ('report.view_calendar',)
    model = ReportPeriod
    template_name = 'calendar-list.html'
    context_object_name = 'reportfile'

    def get_context_data(self, **kwargs):
        country = Country.objects.get(id=self.request.session.get('country'))
        context = super(CalendarList, self).get_context_data(**kwargs)
        years = self.request.GET.getlist('years')
        years.sort()

        context['data'] = ReportPeriod.objects.filter(year__in=years, country=self.request.session.get('country'))
        context['activate'] = ''
        context['q_string'] = self.request.META['QUERY_STRING']
        context['order_by'] = self.request.GET.get('order_by', 'year')
        context['years'] = list(range(2000, datetime.datetime.now().year + 1))
        context['selected_years'] = years
        context['months'] = self.months
        if country.name == 'Canada':
            context['sample_file_path'] = "/media/IATA_calendar_sample.xlsx"
        else:
            context['sample_file_path'] = "/media/ARC_calendar_sample.xlsx"

        return context


class DisbursementSummary(PermissionRequiredMixin, ListView):
    """Disbursement Summary Report"""

    permission_required = ('report.view_disbursement_summary',)
    model = Disbursement
    template_name = 'disbursement-summary-report.html'
    context_object_name = 'disbursementsummary'
    paginate_by = 1000

    def get_queryset(self):
        month_year = self.request.GET.get('month_year', '')
        start_date = self.request.GET.get('start_date', '')
        end_date = self.request.GET.get('end_date', '')
        airline = self.request.GET.get('airline', '')
        trns_list = []

        if airline:
            qs = Disbursement.objects.select_related('report_period').order_by('report_period__ped', 'airline').filter(
                airline=airline)

            if month_year:
                month = datetime.datetime.strptime(month_year, '%B %Y').month or ''
                year = datetime.datetime.strptime(month_year, '%B %Y').year or ''
                if month and year:
                    qs = qs.filter(report_period__month=month, report_period__year=year)

            if start_date and end_date:
                start = datetime.datetime.strptime(start_date, '%d %B %Y')
                end = datetime.datetime.strptime(end_date, '%d %B %Y')
                qs = qs.filter(report_period__ped__range=[start, end])

            qs = qs.annotate(
                totalarc_deductions=F('arc_deduction') + F('arc_fees') + F('arc_reversal') + F('arc_tot')).annotate(
                netw_disbursement=F('bank7') - F('totalarc_deductions'))

            trns_list = []

            for disbursement in qs:
                trn = dict()

                trn['totalarc_deductions'] = disbursement.totalarc_deductions
                trn['totalw_disbursement'] = disbursement.bank7
                trn['netw_disbursement'] = disbursement.netw_disbursement
                trn['ped'] = disbursement.report_period.ped

                trns_list.append(trn)

        return trns_list

    def get_context_data(self, **kwargs):
        context = super(DisbursementSummary, self).get_context_data(**kwargs)

        context['activate'] = 'reports'
        context['month_year'] = self.request.GET.get('month_year', '')
        context['selected_airline'] = self.request.GET.get('airline', '')
        context['date_filter'] = self.request.GET.get('date_filter', 'month_year')
        context['start_date'] = self.request.GET.get('start_date', '')
        context['end_date'] = self.request.GET.get('end_date', '')
        context['airlines'] = Airline.objects.filter(country=self.request.session.get('country'))
        return context


class GetDisbursementSummary(PermissionRequiredMixin, View):
    """Filtered Disbursement Summary Report download as exel."""

    permission_required = ('report.download_commission',)

    def get(self, request, *args, **kwargs):
        month_year = self.request.GET.get('month_year', '')
        start_date = self.request.GET.get('start_date', '')
        end_date = self.request.GET.get('end_date', '')
        airline = self.request.GET.get('airline', '')

        if airline:
            qs = Disbursement.objects.select_related('report_period').filter(airline=airline)

            if month_year:
                month = datetime.datetime.strptime(month_year, '%B %Y').month or ''
                year = datetime.datetime.strptime(month_year, '%B %Y').year or ''
                if month and year:
                    qs = qs.filter(report_period__month=month, report_period__year=year)

            if start_date and end_date:
                start = datetime.datetime.strptime(start_date, '%d %B %Y')
                end = datetime.datetime.strptime(end_date, '%d %B %Y')
                qs = qs.filter(report_period__ped__range=[start, end])

            qs = qs.annotate(
                totalarc_deductions=F('arc_deduction') + F('arc_fees') + F('arc_reversal') + F('arc_tot')).annotate(
                netw_disbursement=F('bank7') - F('totalarc_deductions'))

            response = HttpResponse(content_type='application/vnd.ms-excel')
            airline_obj = Airline.objects.filter(pk=self.request.GET.get('airline', '')).first()
            wb = xlwt.Workbook(encoding='utf-8')
            ws = FitSheetWrapper(wb.add_sheet('Disbursement Summary Report'))
            row_num = 0
            if airline_obj:
                if month_year:
                    dt_rep = month_year

                    ws.row(row_num).height = 20 * 20
                    ws.write_merge(row_num, 0, 0, 7, airline_obj.name.upper(), bold_center)
                    row_num = row_num + 1
                    ws.row(row_num).height = 20 * 20
                    ws.write_merge(row_num, 1, 0, 7, "DISBURSEMENT SUMMARY REPORT", bold_center)
                    row_num = row_num + 1
                    ws.row(row_num).height = 20 * 20
                    ws.write_merge(row_num, 2, 0, 7, month_year.upper(), bold_center)
                    row_num = row_num + 1
                    ws.row(row_num).height_mismatch = True
                else:
                    dt_rep = start_date + " - " + end_date

                    ws.row(row_num).height = 20 * 20
                    ws.write_merge(row_num, 0, 0, 7, "DISBURSEMENT SUMMARY REPORT", bold_center)
                    row_num = row_num + 1
                    ws.row(row_num).height = 20 * 20
                    ws.write_merge(row_num, 1, 0, 7, month_year.upper(), bold_center)
                    row_num = row_num + 1
                    ws.row(row_num).height_mismatch = True

                file_name = airline_obj.abrev + " " + dt_rep + " disbursement_summary_report.xls"
            else:
                file_name = "disbursement_summary.xls"
            response['Content-Disposition'] = 'inline; filename=' + file_name

            peds = qs.values_list('report_period__ped', flat=True).order_by(
                'report_period__ped').distinct()

            for ped in peds:
                ws.write(row_num, 0, "PED: " + str(ped), yellow_background_header)
                row_num = row_num + 2
                columns = ['Total Weekly Cash Disbursement', 'ARC Deductions', 'Net Weekly Cash Disbursement']
                for col_num in range(len(columns)):
                    ws.write(row_num, col_num, columns[col_num], yellow_background_header)

                trns_list = []

                for disbursement in qs.filter(report_period__ped=ped):
                    trn = dict()
                    trn['totalarc_deductions'] = disbursement.totalarc_deductions
                    trn['totalw_disbursement'] = disbursement.bank7
                    trn['netw_disbursement'] = disbursement.netw_disbursement
                    trns_list.append(trn)

                for item in trns_list:
                    row_num = row_num + 1
                    ws.write(row_num, 0, item.get('totalw_disbursement'))
                    ws.write(row_num, 1, item.get('totalarc_deductions'))
                    ws.write(row_num, 2, item.get('netw_disbursement'))

                row_num = row_num + 2

                # Sheet body, remaining rows

            wb.save(response)
            return response
        return HttpResponse('something went wrong')


from django.utils.decorators import method_decorator
from django.views.decorators.csrf import csrf_exempt
@method_decorator(csrf_exempt, name='dispatch')
class SchedulerReportUpload(View):
    """
    Reports downloaded from hosts which is added by user and uploaded to each airlines in a specific interval.
    """
    def post(self, request, *args, **kwargs):
        import logging
        file = request.FILES['file']
        file_r = self.request.FILES['file'].read()
        try:
            filenameee = file.name
            logging.info("FILE NAME      >>>>>        "+str(filenameee))
            logging.info("FILE NAME   ---   >>>>>        "+str(file))
        except Exception as e:
            logging.info("FILE NAME errorrrrr     >>>>>        "+str(e))
            logging.info("FILE NAME   =======   >>>>>        "+str(file))
        logging.info(str(file.name.split('.')))
        file_name, extention = file.name.split('.')
        media_root = getattr(settings, 'MEDIA_ROOT')
        error = None
        error_status = 0
        if self.request.POST.get("from_scheduler") is not None:
            country = Country.objects.get(id=self.request.POST.get("countrycode"))
        else:
            country = Country.objects.get(id=self.request.session.get('country'))

        # Zip file for both IATA and ARC
        if extention.lower() == 'zip':
            errfiles = []
            errMsg = ''
            zipfileCnt = 0

            media_root = getattr(settings, 'MEDIA_ROOT')
            dt = datetime.datetime.now()
            timestamp = int(time.mktime(dt.timetuple()))
            fs = FileSystemStorage(location=os.path.join(media_root, "reportfile/"))
            zipFldr = file_name + '_' + str(timestamp)
            filename = fs.save(zipFldr + '.' + extention, file)

            filenw = ZipFile(os.path.join(media_root, "reportfile/") + filename, 'r')

            for name in filenw.namelist():
                try:
                    if name.endswith('/'):
                        pass
                    else:
                        extractfile_name, extractfile_extension = name.split('.')

                        if extractfile_extension.lower() == 'pdf' and country.name != 'United States':
                            # pdf file for IATA

                            lf = tempfile.NamedTemporaryFile(dir=os.path.join(media_root, "reportfile/"), suffix='.pdf')
                            lf.write(filenw.read(name))
                            text_file = os.path.join(media_root, "reportfile/" + extractfile_name + ".txt")
                            if str(extractfile_name).startswith(country.code + '_FCAIBILLDET'):
                                call(['pdftotext', '-layout', lf.name, text_file])
                                error = process_billing_details(text_file, self.request)
                            elif str(extractfile_name).startswith(country.code + '_PCAIDLYDET'):
                                call(['pdftotext', '-layout', lf.name, text_file])
                                error = process_card_details(text_file, self.request)
                            else:
                                error = "Incorrect file"

                            if error:
                                errfiles.append(name)
                                context = {
                                    'request': self.request,
                                    'error': error,
                                    'file_name': name,
                                    'country': country.name
                                }

                                admin_emails = User.objects.filter(is_superuser=True).values_list('email',
                                                                                                  flat=True)  # super user emails
                                # send_mail("File upload parsing issue.", "email/parsing-issue-email.html", context,
                                #           admin_emails,
                                #           from_email='assda@assda.com')
                            else:
                                context = {
                                    'request': self.request,
                                    'file_name': name,
                                    'country': country.name
                                }
                                admin_emails = User.objects.filter(is_superuser=True).values_list('email',
                                                                                                  flat=True)  # super user emails
                                # send_mail("File uploaded successfully.", "email/success_mail_after_upload.html",
                                #           context,
                                #           admin_emails,
                                #           from_email='Assda@assda.com')
                            # remove tmp file
                            lf.close()

                            # file count
                            zipfileCnt = zipfileCnt + 1

                        elif extractfile_extension.lower() == 'txt' and country.name == 'United States':


                            filenw.extract(name, os.path.join(media_root, "reportfile/"))
                            file_name = extractfile_name

                            oldname = os.path.join(os.path.join(media_root, "reportfile/"), name)
                            stored_filename = file_name + datetime.datetime.now().strftime('%s') + '.txt'

                            filepath = os.path.join(os.path.join(media_root, "reportfile/"), stored_filename)

                            os.rename(oldname, filepath)

                            if str(file_name).startswith('CARRPTSW'):
                                error = process_carrier_report(filepath, self.request)
                            elif str(file_name).startswith('DISBADV'):
                                error = process_disbursement_advice(filepath, file_name, self.request)
                            elif str(file_name).startswith('CARRDED'):
                                error = process_carrier_deductions(filepath, file_name, self.request)
                            else:
                                error = "Incorrect file"
                            logging.info("ERROR : "+str(file_name) +"   ==>>   "+error)
                            if error:
                                
                                if os.path.exists(filepath):
                                    os.remove(filepath)
                                errfiles.append(name)
                                context = {
                                    'request': self.request,
                                    'error': error,
                                    'file_name': name,
                                    'country': country.name
                                }
                                admin_emails = User.objects.filter(is_superuser=True).values_list('email',
                                                                                                  flat=True)  # super user emails
                                # send_mail("File upload parsing issue.", "email/parsing-issue-email.html", context,
                                #           admin_emails,
                                #           from_email='assda@assda.com')
                            else:
                                context = {
                                    'request': self.request,
                                    'file_name': name,
                                    'country': country.name
                                }
                                admin_emails = User.objects.filter(is_superuser=True).values_list('email',
                                                                                                  flat=True)  # super user emails
                                # send_mail("File uploaded successfully.", "email/success_mail_after_upload.html",
                                #           context,
                                #           admin_emails,
                                #           from_email='Assda@assda.com')


                            # file count
                            zipfileCnt = zipfileCnt + 1
                except:
                    pass
            # Remove zip file
            if os.path.exists(os.path.join(media_root, "reportfile/") + filename):
                os.remove(os.path.join(media_root, "reportfile/") + filename)


        elif extention.lower() == 'pdf' and country.name != 'United States':
            # pdf file for IATA
            text_file = os.path.join(media_root, "reportfile/" + file_name + ".txt")

            lf = tempfile.NamedTemporaryFile(dir=os.path.join(media_root, "reportfile/"), suffix='.pdf')
            lf.write(file_r)
            if str(file_name).startswith(country.code + '_FCAIBILLDET'):
                call(['pdftotext', '-layout', lf.name, text_file])
                with open(text_file, 'r+', encoding="utf-8") as fd:
                    lines = fd.readlines()
                    fd.seek(0)
                    fd.writelines(line for line in lines if line.strip())
                    fd.truncate()
                error = process_billing_details(text_file, self.request)
            elif str(file_name).startswith(country.code + '_PCAIDLYDET'):
                call(['pdftotext', '-layout', lf.name, text_file])
                error = process_card_details(text_file, self.request)
            else:
                error = "Incorrect file"

            # remove tmp file
            lf.close()

            if error:
                if not isinstance(error, dict):
                    # form.add_error('file', error)
                    messages.add_message(self.request, messages.ERROR, error)
                else:
                    self.miss_match_errors = error
                    messages.add_message(self.request, messages.ERROR, 'Mismatch in parsed amounts.')
                    context = {
                        'request': self.request,
                        'error': error,
                        'file_name': file.name
                    }
                    admin_emails = list(User.objects.filter(is_superuser=True).values_list('email', flat=True))
                    # admin_emails.append(self.request.user.email)
                    # admin_emails = ['ajithanil7262@gmail.com']

                    # send_mail("File upload parsing issue.", "email/parsing-issue-email.html", context, admin_emails,
                    #           from_email='Assda@assda.com')

                error_status = 1

            else:
                context = {
                    'request': self.request,
                    'file_name': file.name,
                    'country': country.name
                }
                admin_emails = list(User.objects.filter(is_superuser=True).values_list('email', flat=True))
                # admin_emails = ['ajithanil7262@gmail.com']

                # send_mail("File uploaded successfully.", "email/success_mail_after_upload.html", context,
                #           admin_emails,
                #           from_email='Assda@assda.com')
                error_status = 1
                # return super(ReportUpload, self).form_valid(form)
                return HttpResponse("File - {} uploaded successfully".format(file.name))
        elif extention.lower() == 'txt' and country.name == 'United States':
            # pdf file for ARC

            stored_filename = file_name + datetime.datetime.now().strftime('%s') + '.txt'
            filepath = os.path.join(os.path.join(media_root, "reportfile/"), stored_filename)
            filesave = open(filepath, 'wb+')
            for chunk in file.chunks():
                filesave.write(chunk)
            filesave.close()

            if str(file_name).startswith('CARRPTSW'):
                error = process_carrier_report(filepath, self.request)
            elif str(file_name).startswith('DISBADV'):
                error = process_disbursement_advice(filepath, file_name, self.request)
            elif str(file_name).startswith('CARRDED'):
                error = process_carrier_deductions(filepath, file_name, self.request)
            else:
                error = "Incorrect file"
            # ARC

            if error:
                messages.add_message(self.request, messages.ERROR, error)

                context = {
                    'request': self.request,
                    'error': error,
                    'file_name': file.name
                }
                admin_emails = list(User.objects.filter(is_superuser=True).values_list('email', flat=True))


                # send_mail("File upload parsing issue.", "email/parsing-issue-email.html", context, admin_emails,
                #           from_email='Assda@assda.com')
                error_status = 1
            else:
                context = {
                    'request': self.request,
                    'file_name': file.name,
                    'country': country.name
                }
                admin_emails = list(User.objects.filter(is_superuser=True).values_list('email', flat=True))

                # send_mail("File uploaded successfully.", "email/success_mail_after_upload.html", context,
                #           admin_emails,
                #           from_email='Assda@assda.com')

        else:
            # form.add_error('file', 'Unsupported file format.')
            if country.name != 'United States':
                # messages.add_message(self.request, messages.ERROR,
                #                      'Unsupported file format. Please upload zip file/pdf document.')
                context = {
                    'request': self.request,
                    'file_name': file.name,
                    'error': error
                }
                admin_emails = list(User.objects.filter(is_superuser=True).values_list('email', flat=True))


                # send_mail("Unsupported file format. Please upload zip file/txt document.",
                #           "email/parsing-issue-email.html", context,
                #           admin_emails,
                #           from_email='Assda@assda.com')

            else:

                context = {
                    'request': self.request,
                    'file_name': file.name,
                    'error': error
                }
                admin_emails = list(User.objects.filter(is_superuser=True).values_list('email', flat=True))

                # send_mail("Unsupported file format. Please upload zip file/txt document.",
                #           "email/parsing-issue-email.html", context,
                #           admin_emails,
                #           from_email='Assda@assda.com')
                error_status = 1

        return HttpResponse("File - {} uploaded successfully".format(file.name))


class WeeklySalesReport(PermissionRequiredMixin, View):
    """Weekly Airlines Sales Summary Report."""

    permission_required = ('report.view_weekly_sales_report',)
    model = Transaction
    template_name = 'weekly-sales-report.html'
    context_object_name = 'transactions'

    def get(self, request):
        context = dict()
        month_year = self.request.GET.get('month_year', '')
        airline_id = self.request.GET.get('airline', '')
        week_num = self.request.GET.get('week_num', 1)
        all_transactions = Transaction.objects.all()
        qs = Transaction.objects.select_related('report__report_period', 'report__airline')
        airlines = Airline.objects.filter(country=self.request.session.get('country'))
        values = []
        is_arc_var = is_arc(self.request.session.get('country'))
        if is_arc_var:
            qs_disp = Disbursement.objects.select_related('report_period', 'airline')
        else:
            qs_acm = Transaction.objects.select_related('report__report_period', 'report__airline', 'agency').filter(
                transaction_type__startswith='ACM', agency__agency_no='6998051')
        has_transaction = True
        if month_year and airline_id:
            airline = Airline.objects.get(id=airline_id)
            month = datetime.datetime.strptime(month_year, '%B %Y').month or ''
            year = datetime.datetime.strptime(month_year, '%B %Y').year or ''
            context['month_year2'] = datetime.datetime(year, month, 1).strftime("%b %Y")
            # context['month_year1'] = (datetime.datetime(year, month, 1) - rd(years=1)).strftime("%b %Y")

            if month and year:
                # for week_num in [1,2,3,4,5]:
                # qs1 = qs.filter(report__airline=airline,report__report_period__month=month, report__report_period__year=year - 1)
                qs2 = qs.filter(report__airline=airline, report__report_period__month=month,
                                report__report_period__year=year)

                if not is_arc_var:
                    qs2_acm = qs_acm.filter(report__airline=airline, report__report_period__month=month,
                                            report__report_period__year=year)

                sales_data = {}
                sales_data['airline'] = airline
                if qs2:
                    if is_arc_var:


                        sub_tax_cp_mf = Charges.objects.select_related('transaction').filter(type__in=['CP', 'MF'],
                                                                                             transaction__report__report_period__ped=OuterRef(
                                                                                                 'report__report_period__ped'),
                                                                                             transaction__report__airline=airline).values(

                            'transaction__report__report_period__ped').annotate(
                            dcount=Count('transaction__report__report_period__ped'),
                            total_tax_cp_mf=Sum('amount', output_field=FloatField())).values('total_tax_cp_mf')
                        t_filtr = Transaction.objects.exclude(
                            transaction_type__startswith='ACM',
                            agency__agency_no__in=['6998051', ]).select_related('report',
                                                                                'report__airline',
                                                                                'report__report_period').filter(
                            report__airline=airline,
                            report__report_period__month=month,
                            report__report_period__year=year, is_sale=True)

                        transactions = Transaction.objects.exclude(
                            transaction_type__startswith='ACM',
                            agency__agency_no__in=['6998051', ]).select_related(
                            'report',
                            'report__airline',
                            'report__report_period').filter(
                            report__airline=airline,
                            report__report_period__month=month,
                            report__report_period__year=year, is_sale=True).values(
                            'report__report_period__ped', 'report__report_period__week', 'card_code',
                            'report__report_period__remittance_date').annotate(
                            dcount=Count('card_code'),
                            total_fare_amount=Sum('fare_amount'),
                            total_commission=Sum('std_comm_amount'),
                                        pen_value=Sum('pen'),
                            cp=Sum('pen', filter=Q(report__airline=airline, pen_type='CANCEL PEN')),

                            total_ap_acm=Sum('transaction_amount',
                                             filter=Q(report__airline=airline, transaction_type__startswith='ACM',
                                                      agency__agency_no='31768026')),

                        ).order_by(
                            'report__report_period__ped')

                        report_period_list = list(
                            transactions.values_list('report__report_period__ped', flat=True).distinct())


                        cash_fare_total = 0
                        amex_fare_total = 0
                        visa_fare_total = 0
                        mastercard_fare_total = 0
                        other_fare_total = 0

                        cash_comm_total = 0
                        amex_comm_total = 0
                        visa_comm_total = 0
                        mastercard_comm_total = 0
                        other_comm_total = 0

                        cash_tr_total = 0
                        amex_tr_total = 0
                        visa_tr_total = 0
                        mastercard_tr_total = 0
                        other_tr_total = 0

                        cash_cp_total = 0
                        amex_cp_total = 0
                        visa_cp_total = 0
                        mastercard_cp_total = 0
                        other_cp_total = 0
                        weekly_sales_value_toal = 0

                        c_total = 0
                        v_total = 0
                        m_total = 0
                        a_total = 0
                        o_total = 0

                        cash_tax_total = 0
                        amex_tax_total = 0
                        visa_tax_total = 0
                        mastercard_tax_total = 0
                        other_tax_total = 0

                        cash_yqyr_total = 0
                        amex_yqyr_total = 0
                        visa_yqyr_total = 0
                        mastercard_yqyr_total = 0
                        other_yqyr_total = 0

                        num = 0
                        for report_period in report_period_list:
                            cash_transaction_amount = t_filtr.filter(
                                Q(transaction_type__startswith='ACM', agency__agency_no='31768026')).filter(
                                Q(card_code="") | Q(card_code=None),
                                report__report_period__ped=report_period).aggregate(
                                total_tr_value=Sum('transaction_amount')).get("total_tr_value")
                            amex_transaction_amount = t_filtr.filter(
                                Q(transaction_type__startswith='ACM', agency__agency_no='31768026')).filter(
                                card_code__startswith="3", report__report_period__ped=report_period).aggregate(
                                total_tr_value=Sum('transaction_amount')).get("total_tr_value")
                            visa_transaction_amount = t_filtr.filter(
                                Q(transaction_type__startswith='ACM', agency__agency_no='31768026')).filter(
                                card_code__startswith="4", report__report_period__ped=report_period).aggregate(
                                total_tr_value=Sum('transaction_amount')).get("total_tr_value")
                            mastercard_transaction_amount = t_filtr.filter(
                                Q(transaction_type__startswith='ACM', agency__agency_no='31768026')).filter(
                                card_code__startswith="5", report__report_period__ped=report_period).aggregate(
                                total_tr_value=Sum('transaction_amount')).get("total_tr_value")
                            other_transaction_amount = t_filtr.filter(
                                Q(transaction_type__startswith='ACM', agency__agency_no='31768026')).filter(
                                ~Q(card_code=None) & ~Q(card_code="") & ~Q(card_code="3") & ~Q(card_code="4") & ~Q(
                                    card_code="5"), report__report_period__ped=report_period).aggregate(
                                total_tr_value=Sum('transaction_amount')).get("total_tr_value")

                            if not cash_transaction_amount:
                                cash_transaction_amount = 0
                            if not amex_transaction_amount:
                                amex_transaction_amount = 0
                            if not visa_transaction_amount:
                                visa_transaction_amount = 0
                            if not mastercard_transaction_amount:
                                mastercard_transaction_amount = 0
                            if not other_transaction_amount:
                                other_transaction_amount = 0



                            cash_pen_amount = t_filtr.filter(
                                Q(card_code="") | Q(card_code=None),
                                report__report_period__ped=report_period).aggregate(
                                total_comm_value=Sum('pen')).get("total_comm_value")
                            amex_pen_amount = t_filtr.filter(
                                card_code__startswith="3", report__report_period__ped=report_period).aggregate(
                                total_comm_value=Sum('pen')).get("total_comm_value")
                            visa_pen_amount = t_filtr.filter(
                                card_code__startswith="4", report__report_period__ped=report_period).aggregate(
                                total_comm_value=Sum('pen')).get("total_comm_value")
                            mastercard_pen_amount = t_filtr.filter(
                                card_code__startswith="5", report__report_period__ped=report_period).aggregate(
                                total_comm_value=Sum('pen')).get("total_comm_value")
                            other_pen_amount = t_filtr.filter(
                                ~Q(card_code=None) & ~Q(card_code="") & ~Q(card_code__startswith="3") & ~Q(
                                    card_code__startswith="4") & ~Q(
                                    card_code__startswith="5"), report__report_period__ped=report_period).aggregate(
                                total_comm_value=Sum('pen')).get("total_comm_value")

                            amex_yq_yr_amount = Charges.objects.select_related('transaction').filter(
                                type__in=['YQ', 'YR'], transaction__report__report_period__ped=report_period,
                                transaction__report__airline=airline, transaction__card_code__startswith="3").aggregate(
                                total_comm_value=Sum('amount')).get("total_yqyr_value")
                            visa_yq_yr_amount = Charges.objects.select_related('transaction').filter(
                                type__in=['YQ', 'YR'], transaction__report__report_period__ped=report_period,
                                transaction__report__airline=airline, transaction__card_code__startswith="4").aggregate(
                                total_comm_value=Sum('amount')).get("total_yqyr_value")
                            mastercard_yq_yr_amount = Charges.objects.select_related('transaction').filter(
                                type__in=['YQ', 'YR'], transaction__report__report_period__ped=report_period,
                                transaction__report__airline=airline, transaction__card_code__startswith="5").aggregate(
                                total_comm_value=Sum('amount')).get("total_yqyr_value")
                            cash_yq_yr_amount = Charges.objects.select_related('transaction').filter(
                                Q(transaction__card_code=None) | Q(transaction__card_code=""), type__in=['YQ', 'YR'],
                                transaction__report__report_period__ped=report_period,
                                transaction__report__airline=airline).aggregate(total_comm_value=Sum('amount')).get(
                                "total_yqyr_value")
                            other_yq_yr_amount = Charges.objects.select_related('transaction').filter(
                                ~Q(transaction__card_code=None) & ~Q(transaction__card_code="") & ~Q(
                                    transaction__card_code__startswith="3") & ~Q(
                                    transaction__card_code__startswith="4") & ~Q(
                                    transaction__card_code__startswith="5"), type__in=['YQ', 'YR'],
                                transaction__report__report_period__ped=report_period).aggregate(
                                total_comm_value=Sum('amount')).get("total_yqyr_value")

                            amex_tax_amount = Taxes.objects.select_related('transaction').filter(
                                transaction__report__report_period__ped=report_period,
                                transaction__report__airline=airline, transaction__card_code__startswith="3").aggregate(
                                total_comm_value=Sum('amount')).get("total_comm_value")
                            visa_tax_amount = Taxes.objects.select_related('transaction').filter(
                                transaction__report__report_period__ped=report_period,
                                transaction__report__airline=airline, transaction__card_code__startswith="4").aggregate(
                                total_comm_value=Sum('amount')).get("total_comm_value")
                            mastercard_tax_amount = Taxes.objects.select_related('transaction').filter(
                                transaction__report__report_period__ped=report_period,
                                transaction__report__airline=airline, transaction__card_code__startswith="5").aggregate(
                                total_comm_value=Sum('amount')).get("total_comm_value")
                            cash_tax_amount = Taxes.objects.select_related('transaction').filter(
                                Q(transaction__card_code=None) | Q(transaction__card_code=""),
                                transaction__report__report_period__ped=report_period,
                                transaction__report__airline=airline).aggregate(total_comm_value=Sum('amount')).get(
                                "total_comm_value")
                            other_tax_amount = Taxes.objects.select_related('transaction').filter(
                                ~Q(transaction__card_code=None) & ~Q(transaction__card_code="") & ~Q(
                                    transaction__card_code__startswith="3") & ~Q(
                                    transaction__card_code__startswith="4") & ~Q(
                                    transaction__card_code__startswith="5"),
                                transaction__report__report_period__ped=report_period,
                                transaction__report__airline=airline).aggregate(total_comm_value=Sum('amount')).get(
                                "total_comm_value")

                            if not cash_pen_amount:
                                cash_pen_amount = 0
                            if not amex_pen_amount:
                                amex_pen_amount = 0
                            if not visa_pen_amount:
                                visa_pen_amount = 0
                            if not mastercard_pen_amount:
                                mastercard_pen_amount = 0
                            if not other_pen_amount:
                                other_pen_amount = 0

                            if not amex_yq_yr_amount:
                                amex_yq_yr_amount = 0
                            if not visa_yq_yr_amount:
                                visa_yq_yr_amount = 0
                            if not mastercard_yq_yr_amount:
                                mastercard_yq_yr_amount = 0
                            if not cash_yq_yr_amount:
                                cash_yq_yr_amount = 0
                            if not other_yq_yr_amount:
                                other_yq_yr_amount = 0

                            if not amex_tax_amount:
                                amex_tax_amount = 0
                            if not visa_tax_amount:
                                visa_tax_amount = 0
                            if not mastercard_tax_amount:
                                mastercard_tax_amount = 0
                            if not cash_tax_amount:
                                cash_tax_amount = 0
                            if not other_tax_amount:
                                other_tax_amount = 0



                            cash_cp_total += cash_pen_amount
                            amex_cp_total += amex_pen_amount
                            visa_cp_total += visa_pen_amount
                            mastercard_cp_total += mastercard_pen_amount
                            other_cp_total += other_pen_amount

                            cash_tr_total += cash_transaction_amount
                            amex_tr_total += amex_transaction_amount
                            visa_tr_total += visa_transaction_amount
                            mastercard_tr_total += mastercard_transaction_amount
                            other_tr_total += other_transaction_amount


                            cash_tax_total += cash_tax_amount
                            amex_tax_total += amex_tax_amount
                            visa_tax_total += visa_tax_amount
                            mastercard_tax_total += mastercard_tax_amount
                            other_tax_total += other_tax_amount

                            cash_yqyr_total += cash_yq_yr_amount
                            amex_yqyr_total += amex_yq_yr_amount
                            visa_yqyr_total += visa_yq_yr_amount
                            mastercard_yqyr_total += mastercard_yq_yr_amount
                            other_yqyr_total += other_yq_yr_amount


                        cash_fare_total = 0
                        amex_fare_total = 0
                        visa_fare_total = 0
                        mastercard_fare_total = 0
                        other_fare_total = 0
                        qs = Transaction.objects.select_related('agency').prefetch_related('taxes_set').filter(
                            is_sale=True).annotate(
                            yq=Sum('charges__amount', filter=Q(charges__type='YQ')),
                            yr=Sum('charges__amount', filter=Q(charges__type='YR')))
                        qs = qs.filter(report__airline=airline)
                        qs = qs.filter(report__report_period__month=month, report__report_period__year=year)

                        peds = qs.values_list('report__report_period__ped', flat=True).order_by(
                            'report__report_period__ped').distinct()

                        # total_cashhh = 0
                        for ped in peds:
                            for item in qs.filter(report__report_period__ped=ped):
                                if str(item.card_code) == "" or str(item.card_code) == "None":
                                    cash_fare_total += float(item.fare_amount) if item.fare_amount else 0
                                    cash_comm_total += float(item.std_comm_amount) if item.std_comm_amount else 0


                                elif str(item.card_code)[0] == "3":
                                    if item.fare_amount:
                                        amex_fare_total += float(item.fare_amount) if item.fare_amount else 0
                                        amex_comm_total += float(item.std_comm_amount) if item.std_comm_amount else 0


                                elif str(item.card_code)[0] == "4":
                                    if item.fare_amount:
                                        visa_fare_total += float(item.fare_amount) if item.fare_amount else 0
                                        visa_comm_total += float(item.std_comm_amount) if item.std_comm_amount else 0


                                elif str(item.card_code)[0] == "5":
                                    if item.fare_amount:
                                        mastercard_fare_total += float(item.fare_amount) if item.fare_amount else 0
                                        mastercard_comm_total += float(
                                            item.std_comm_amount) if item.std_comm_amount else 0


                                else:
                                    if item.fare_amount:
                                        other_fare_total += float(item.fare_amount) if item.fare_amount else 0
                                        other_comm_total += float(item.std_comm_amount) if item.std_comm_amount else 0

                        cash_fare_total = cash_fare_total + cash_cp_total
                        amex_fare_total = amex_fare_total + amex_cp_total
                        visa_fare_total = visa_fare_total + visa_cp_total
                        mastercard_fare_total = mastercard_fare_total + mastercard_cp_total
                        other_fare_total = other_fare_total + other_cp_total

                        cash_weekly_sales = (cash_fare_total + cash_tax_total) - cash_comm_total
                        amex_weekly_sales = (amex_fare_total + amex_tax_total) - amex_comm_total
                        visa_weekly_sales = (visa_fare_total + visa_tax_total) - visa_comm_total
                        mastercard_weekly_sales = (mastercard_fare_total + mastercard_tax_total) - mastercard_comm_total
                        other_weekly_sales = (other_fare_total + other_tax_total) - other_comm_total

                        transactions2 = t_filtr.values(
                            # 'report__report_period__ped', 'report__report_period__week', 'report__cc', 'report__ca').annotate(
                            'report__report_period__ped', 'report__report_period__week',
                            'report__report_period__remittance_date').annotate(
                            dcount=Count('report__report_period__ped'),
                            total_fare_amount=Sum('fare_amount'),
                            total_commission=Sum('std_comm_amount'),
                            # report__cc=Sum('cc'),
                            # report__ca=Sum('ca'),
                            pen_value=Sum('pen'),
                            cp=Sum('pen', filter=Q(report__airline=airline, pen_type='CANCEL PEN')),

                            total_ap_acm=Sum('transaction_amount',
                                             filter=Q(report__airline=airline, transaction_type__startswith='ACM',
                                                      agency__agency_no='31768026')),
                            total_tax_cp_mf=Subquery(sub_tax_cp_mf, output_field=FloatField())

                        ).order_by(
                            'report__report_period__ped')
                        arc_deductions_total = 0

                        sales_data['tax_amount'] = 0.0
                        sales_data['fare_amount'] = 0.0
                        sales_data['comm_amount'] = 0.0
                        # for tt in transactions:
                        total_acm_amount = 0
                        total_cash_disbursement = 0
                        total_card_disbursement = 0
                        if transactions and transactions2:
                            visa = 0
                            visa_tax = 0
                            amex = 0
                            amex_tax = 0
                            mastercard = 0
                            mastercard_tax = 0
                            cash = 0
                            cash_tax = 0
                            mastercard_comm = 0
                            visa_comm = 0
                            amex_comm = 0
                            cash_comm = 0
                            mastercard_cpmf = 0
                            visa_cpmf = 0
                            amex_cpmf = 0
                            cash_cpmf = 0
                            mastercard_acm_fare = 0
                            visa_acm_fare = 0
                            amex_acm_fare = 0
                            cash_acm_fare = 0
                            other = 0
                            other_comm = 0
                            other_tax = 0
                            other_cpmf = 0
                            other_acm_fare = 0
                            total_ap_acm_var = 0
                            total_tax_yq_yr_total = 0

                            cash_tax_yq_yr = 0
                            amex_tax_yq_yr = 0
                            visa_tax_yq_yr = 0
                            mastercard_tax_yq_yr = 0
                            other_tax_yq_yr = 0

                            cash_total_acm_amount = 0
                            amex_total_acm_amount = 0
                            visa_total_acm_amount = 0
                            mastercard_total_acm_amount = 0
                            other_total_acm_amount = 0
                            weekly_total_sales = 0
                            # total_fare_amount = 0
                            # total_comm_amount = 0
                            total_tax_amount = 0
                            weekly_total_cash_disbursement_total = 0
                            cash_disbursement_total = 0
                            total_ap_acm = 0
                            test1 = 0
                            test2 = 0
                            test3 = 0
                            test4 = 0
                            total_weekly_sales_total = 0

                            for t_obj in transactions2:

                                total_tax_yq_yr = Charges.objects.select_related('transaction').filter(
                                    type__in=['YQ', 'YR'],
                                    transaction__report__report_period__ped=t_obj.get('report__report_period__ped'),
                                    transaction__report__airline=airline).aggregate(
                                    total_tax_yq_yr=Coalesce(Sum('amount'), V(0))).get('total_tax_yq_yr')

                                if t_obj.get('total_fare_amount'):

                                    total_fare_amount_var = t_obj.get('total_fare_amount')
                                else:
                                    total_fare_amount_var = 0.00

                                if t_obj.get('cp'):
                                    cp_var = t_obj.get('cp')
                                else:
                                    cp_var = 0.00
                                weekly_total_sales += total_fare_amount_var
                                total_fare_amount_var = total_fare_amount_var + cp_var
                                test1 += total_fare_amount_var
                                total_ap_acm_var = 0.00
                                if t_obj.get('total_ap_acm'):
                                    total_ap_acm_var = t_obj.get('total_ap_acm')
                                    total_ap_acm += total_ap_acm_var
                                test2 += total_ap_acm_var

                                total_fare_amount_var = total_fare_amount_var - total_ap_acm_var
                                # net_sales_weekly = total_fare_amount_var
                                test3 += total_fare_amount_var
                                total_tax_var = Taxes.objects.select_related('transaction').filter(
                                    transaction__report__report_period__ped=t_obj.get('report__report_period__ped'),
                                    transaction__report__airline=airline).aggregate(
                                    total_tax=Coalesce(Sum('amount'), V(0))).get(
                                    'total_tax')


                                if t_obj.get('total_commission'):
                                    total_commission_var = t_obj.get('total_commission')
                                else:
                                    total_commission_var = 0.00
                                weekly_sales = total_fare_amount_var + total_tax_var + total_commission_var
                                total_weekly_sales_total += weekly_sales
                                try:
                                    remittance = t.get('report__report_period__remittance_date')
                                except:
                                    pass

                                try:
                                    disb = Disbursement.objects.get(
                                        report_period__ped=t_obj.get('report__report_period__ped'),
                                        airline=airline)
                                    cash_disbursement = disb.bank7
                                    arc_deductions = disb.arc_deduction + disb.arc_fees + disb.arc_reversal + disb.arc_tot
                                    weekly_total_cash_disbursement = disb.arc_net_disb
                                except Exception as e:
                                    cash_disbursement = 0.00
                                    arc_deductions = 0.00
                                    weekly_total_cash_disbursement = 0.00
                                    # pending_deduction = 0.00



                                arc_deductions_total = arc_deductions_total + arc_deductions

                                cash_disbursement_total += cash_disbursement
                                # if weekly_total_cash_disbursement > 0:
                                weekly_total_cash_disbursement_total = weekly_total_cash_disbursement_total + weekly_total_cash_disbursement



                            ii = 0


                            total_weekly_sales = cash_weekly_sales + amex_weekly_sales + visa_weekly_sales + mastercard_weekly_sales + other_weekly_sales
                            total_card_disbursement = total_weekly_sales - cash_disbursement_total
                            cash_data = {
                                'amount': cash_fare_total,
                                'tax': cash_tax_total,
                                'comm_amount': cash_comm_total,
                                'net_sales': cash_weekly_sales
                            }
                            amex_data = {
                                'amount': amex_fare_total,
                                'tax': amex_tax_total,
                                'comm_amount': amex_comm_total,
                                'net_sales': amex_weekly_sales
                            }
                            visa_data = {
                                'amount': visa_fare_total,
                                'tax': visa_tax_total,
                                'comm_amount': visa_comm_total,
                                'net_sales': visa_weekly_sales
                            }
                            mastercard_data = {
                                'amount': mastercard_fare_total,
                                'tax': mastercard_tax_total,
                                'comm_amount': mastercard_comm_total,
                                'net_sales': mastercard_weekly_sales
                            }
                            other_data = {
                                'amount': other_fare_total,
                                'tax': other_tax_total,
                                'comm_amount': other_comm_total,
                                'net_sales': other_weekly_sales
                            }
                            card_data = {
                                'cash_data': cash_data,
                                'amex_data': amex_data,
                                'visa_data': visa_data,
                                'mastercard_data': mastercard_data,
                                'other_data': other_data
                            }
                            total_data = {
                                'total_fare': cash_fare_total + amex_fare_total + visa_fare_total + mastercard_fare_total + other_fare_total,
                                'total_tax': cash_tax_total + amex_tax_total + visa_tax_total + mastercard_tax_total + other_tax_total,
                                'total_comm': cash_comm_total + amex_comm_total + visa_comm_total + mastercard_comm_total + other_comm_total,
                                'total_net_sales': total_weekly_sales,
                                'total_acm_amount': total_ap_acm,
                                'arc_deductions_total': arc_deductions_total,
                                'total_cash_disbursement': weekly_total_cash_disbursement_total,
                                'total_card_disbursement': total_card_disbursement,
                            }
                            sales_data['card_data'] = card_data
                            sales_data['total_data'] = total_data


                    else:


                        qs2_acm_nw = qs2_acm.filter(report__airline=airline).aggregate(
                            total=Sum('fare_amount', output_field=FloatField()),
                            total_comm=Sum('std_comm_amount', output_field=FloatField()),
                            total_cp=Sum('pen', output_field=FloatField()))

                        sales_data['net_sales'] = qs2.filter(report__airline=airline).aggregate(
                            amount2=(Sum('fare_amount') - Coalesce(qs2_acm_nw.get('total'), V(0))) - (
                                Sum('std_comm_amount', output_field=FloatField()) - Coalesce(
                                qs2_acm_nw.get('total_comm'), V(0))) + (
                                        Sum('pen') - Coalesce(qs2_acm_nw.get('total_cp'), V(0)))).get(
                            'amount2') or 0

                        qs2_acm_nw = qs2_acm.filter(report__airline=airline).aggregate(
                            total=Sum('transaction_amount', output_field=FloatField()),
                            total_comm=Sum('std_comm_amount', output_field=FloatField()))


                        sales_data['gross_sales'] = qs2.filter(report__airline=airline).aggregate(
                            amount2=(Sum('transaction_amount', output_field=FloatField()) - Coalesce(
                                qs2_acm_nw.get('total'), V(0))) - (
                                        Sum('std_comm_amount', output_field=FloatField()) - Coalesce(
                                        qs2_acm_nw.get('total_comm'), V(0)))).get(
                            'amount2') or 0

                        sub_tax = Taxes.objects.select_related('transaction').filter(
                            transaction__report__report_period__ped=OuterRef('report__report_period__ped'),
                            transaction__report__airline=airline).values(
                            'transaction__report__report_period__ped').annotate(
                            dcount=Count('transaction__report__report_period__ped'),
                            total_tax=Sum('amount', output_field=FloatField())).values('total_tax')

                        sub_tax_cp_mf = Charges.objects.select_related('transaction').filter(type__in=['CP', 'MF'],
                                                                                             transaction__report__report_period__ped=OuterRef(
                                                                                                 'report__report_period__ped'),
                                                                                             transaction__report__airline=airline).values(

                            'transaction__report__report_period__ped').annotate(
                            dcount=Count('transaction__report__report_period__ped'),
                            total_tax_cp_mf=Sum('amount', output_field=FloatField())).values('total_tax_cp_mf')

                        transactions = all_transactions.filter(
                            report__airline=airline,
                            report__report_period__month=month,
                            report__report_period__year=year).values(
                            # 'report__report_period__ped', 'report__report_period__week', 'report__cc', 'report__ca').annotate(
                            'id', 'card_type', 'fare_amount', 'ticket_no', 'std_comm_amount',
                            'report__report_period__ped', 'report__report_period__week',
                            'report__report_period__remittance_date').annotate(
                            dcount=Count('report__report_period__ped'),
                            total_fare_amount=Sum('fare_amount'),
                            total_commission=Sum('std_comm_amount'),
                            report__cc=Sum('cc'),
                            report__ca=Sum('ca'),
                            pen_value=Sum('pen'),
                            cp=Sum('pen', filter=Q(report__airline=airline, pen_type='CP')),
                            total_tax=Subquery(sub_tax, output_field=FloatField()),

                            total_ap_acm=Sum('transaction_amount',
                                             filter=Q(report__airline=airline, transaction_type__startswith='ACM',
                                                      agency__agency_no='6998051')),
                            total_tax_cp_mf=Subquery(sub_tax_cp_mf, output_field=FloatField())
                        ).order_by(
                            'report__report_period__ped')


                        arc_deductions_total = 0

                        sales_data['tax_amount'] = 0.0
                        sales_data['fare_amount'] = 0.0
                        sales_data['comm_amount'] = 0.0
                        total_acm_amount = 0
                        total_cash_disbursement = 0
                        total_card_disbursement = 0
                        if transactions:
                            for t_obj in transactions:
                                if t_obj.get("total_ap_acm"):
                                    total_acm_amount += t_obj.get("total_ap_acm")

                                try:
                                    disb = Disbursement.objects.get(
                                        report_period__ped=t_obj.get('report__report_period__ped'),
                                        airline=airline)
                                    cash_disbursement = disb.bank7
                                    arc_deductions = disb.arc_deduction + disb.arc_fees + disb.arc_reversal + disb.arc_tot
                                    weekly_total_cash_disbursement = disb.arc_net_disb
                                except Exception as e:
                                    cash_disbursement = 0.00
                                    arc_deductions = 0.00
                                    weekly_total_cash_disbursement = 0.00

                                arc_deductions_total = arc_deductions_total + arc_deductions

                                total_cash_disbursement += t_obj.get("report__ca") - t_obj.get('total_commission')
                                if t_obj.get('report__cc'):
                                    total_card_disbursement += t_obj.get('report__cc')
                            sales_data['tax_amount'] = transactions[0].get('total_tax')
                            sales_data['fare_amount'] = transactions[0].get('total_fare_amount')
                            sales_data['comm_amount'] = transactions[0].get('total_commission')
                            t = transactions[0]
                            if t.get("total_tax_cp_mf"):
                                total_tax_cp_mf = t.get("total_tax_cp_mf")
                            else:
                                total_tax_cp_mf = 0.00

                            ap_acm = all_transactions.filter(
                                report__report_period__ped=t.get('report__report_period__ped'),
                                report__airline=airline, transaction_type__startswith='ACM',
                                agency__agency_no='6998051').aggregate(
                                total=Sum('fare_amount', output_field=FloatField()),
                                total_comm=Sum('std_comm_amount', output_field=FloatField()),
                                total_cp=Sum('pen', output_field=FloatField()))

                            visa = 0
                            visa_tax = 0
                            amex = 0
                            amex_tax = 0
                            mastercard = 0
                            mastercard_tax = 0
                            cash = 0
                            cash_tax = 0
                            mastercard_comm = 0
                            visa_comm = 0
                            amex_comm = 0
                            cash_comm = 0
                            mastercard_cpmf = 0
                            visa_cpmf = 0
                            amex_cpmf = 0
                            cash_cpmf = 0
                            mastercard_acm_fare = 0
                            visa_acm_fare = 0
                            amex_acm_fare = 0
                            cash_acm_fare = 0
                            other = 0
                            other_comm = 0
                            other_tax = 0
                            other_cpmf = 0
                            other_acm_fare = 0
                            for card in transactions:
                                card_type = str(card.get("card_type")).replace(" ", "")
                                taxes = Taxes.objects.filter(transaction__id=card.get("id"))
                                charges = Charges.objects.filter(type__in=['CP', 'MF'], transaction__id=card.get("id"))
                                acm_data = all_transactions.filter(
                                    report__report_period__ped=card.get('report__report_period__ped'),
                                    report__airline=airline, transaction_type__startswith='ACM',
                                    agency__agency_no='6998051', card_type=card.get("card_type"))
                                total_acm_fare = 0
                                total_tax = 0
                                total_charges = 0
                                for tax in taxes:
                                    total_tax += float(tax.amount)
                                for charge in charges:
                                    total_charges += charge.amount
                                for acm in acm_data:
                                    total_acm_fare += acm.fare_amount
                                if card_type == "Mastercard":
                                    mastercard += float(card.get("fare_amount"))
                                    mastercard_comm += float(card.get("std_comm_amount"))
                                    mastercard_tax += total_tax
                                    mastercard_cpmf += total_charges
                                    mastercard_acm_fare += total_acm_fare
                                elif card_type == "VISAInternational":
                                    visa += float(card.get("fare_amount"))
                                    visa_comm += float(card.get("std_comm_amount"))
                                    visa_tax += total_tax
                                    visa_cpmf += total_charges
                                    visa_acm_fare += total_acm_fare
                                elif card_type == "AmericanExpress":
                                    amex += float(card.get("fare_amount"))
                                    amex_comm += float(card.get("std_comm_amount"))
                                    amex_tax += total_tax
                                    amex_cpmf += total_charges
                                    amex_acm_fare += total_acm_fare
                                elif card_type != "None" and card_type not in ['AmericanExpress', 'VISAInternational',
                                                                               'Mastercard']:
                                    other += 0
                                    other_comm += float(card.get("std_comm_amount"))
                                    other_tax += total_tax
                                    other_cpmf += total_charges
                                    other_acm_fare += total_acm_fare
                                else:
                                    cash += float(card.get("fare_amount"))
                                    cash_comm += float(card.get("std_comm_amount"))
                                    cash_tax += total_tax
                                    cash_cpmf += total_charges
                                    cash_acm_fare += total_acm_fare
                            cash_netsales = cash - cash_comm + cash_tax
                            amex_netsales = amex - amex_comm + amex_tax
                            visa_netsales = visa - amex_comm + amex_tax
                            mastercard_netsales = mastercard - amex_comm + amex_tax
                            other_netsales = other - other_comm + other_tax
                            cash_data = {
                                'amount': cash,
                                'tax': cash_tax,
                                'comm_amount': cash_comm,
                                'net_sales': cash_netsales
                            }
                            amex_data = {
                                'amount': amex,
                                'tax': amex_tax,
                                'comm_amount': amex_comm,
                                'net_sales': amex_netsales
                            }
                            visa_data = {
                                'amount': visa,
                                'tax': visa_tax,
                                'comm_amount': visa_comm,
                                'net_sales': visa_netsales
                            }
                            mastercard_data = {
                                'amount': mastercard,
                                'tax': mastercard_tax,
                                'comm_amount': mastercard_comm,
                                'net_sales': mastercard_netsales
                            }
                            other_data = {
                                'amount': other,
                                'tax': other_tax,
                                'comm_amount': other_comm,
                                'net_sales': other_netsales
                            }
                            card_data = {
                                'cash_data': cash_data,
                                'amex_data': amex_data,
                                'visa_data': visa_data,
                                'mastercard_data': mastercard_data,
                                'other_data': other_data
                            }
                            total_data = {
                                'total_fare': cash + amex + visa + mastercard,
                                'total_tax': cash_tax + amex_tax + visa_tax + mastercard_tax,
                                'total_comm': cash_comm + amex_comm + visa_comm + mastercard_comm,
                                'total_net_sales': cash_netsales + amex_netsales + visa_netsales + mastercard_netsales,
                                'total_acm_amount': total_acm_amount,
                                'arc_deductions_total': arc_deductions_total,
                                'total_cash_disbursement': total_cash_disbursement,
                                'total_card_disbursement': total_card_disbursement,
                            }
                            sales_data['card_data'] = card_data
                            sales_data['total_data'] = total_data

                            if t.get('total_fare_amount'):
                                total_ap_acm_var = 0.00
                                if ap_acm.get('total'):
                                    total_ap_acm_var = ap_acm.get('total')
                                total_fare_amount_var = t.get('total_fare_amount') - total_ap_acm_var
                            else:
                                total_fare_amount_var = 0.00
                            if t.get('total_tax'):
                                total_tax_var = t.get('total_tax')
                            else:
                                total_tax_var = 0.00

                            total_commission_var = 0.00
                            if t.get('total_commission'):
                                total_commission_var = t.get('total_commission')

                            ap_acm_comm = 0.00
                            if ap_acm.get('total_comm'):
                                ap_acm_comm = ap_acm.get('total_comm')

                            total_commission_var = total_commission_var - ap_acm_comm
                            net_sales_weekly = total_fare_amount_var - total_commission_var + total_tax_cp_mf

                            sales_data['net_sales'] = net_sales_weekly

                    values.append(sales_data)
            context['values'] = values
        context['activate'] = 'reports'
        context['has_transaction'] = has_transaction
        context['sales_type'] = self.request.GET.get('sales_type', '')
        context['airlines'] = airlines
        context['selected_airline'] = airline_id
        context['month_year'] = self.request.GET.get('month_year', '')
        context['week_num'] = self.request.GET.get('week_num', 1)
        context['values'] = values
        context['is_arc'] = is_arc(self.request.session.get('country'))
        return render(request, self.template_name, context)


class GetWeeklySalesReport(PermissionRequiredMixin, View):
    """Filtered Weekly Sales Report download as exel."""

    permission_required = ('report.download_weekly_sales',)

    def get(self, request, *args, **kwargs):
        month_year = self.request.GET.get('month_year', '')
        airline_id = self.request.GET.get('airline', '')
        week_num = self.request.GET.get('week_num', 1)
        # qs = ReportFile.objects.select_related('report_period', 'airline')
        all_transactions = Transaction.objects.all()
        qs = all_transactions.select_related('report__report_period', 'report__airline')
        airlines = Airline.objects.filter(country=self.request.session.get('country'))
        is_arc_var = is_arc(self.request.session.get('country'))
        values = []
        if is_arc_var:
            qs_disp = Disbursement.objects.select_related('report_period', 'airline')
        else:
            qs_acm = all_transactions.select_related('report__report_period', 'report__airline', 'agency').filter(
                transaction_type__startswith='ACM', agency__agency_no='6998051')


        if month_year and airline_id:
            airline = Airline.objects.get(id=airline_id)
            month = datetime.datetime.strptime(month_year, '%B %Y').month or ''
            year = datetime.datetime.strptime(month_year, '%B %Y').year or ''

            if month and year:
                qs2 = qs.filter(report__airline=airline, report__report_period__month=month,
                                report__report_period__year=year)
                if is_arc_var:
                    qs1_disp = qs_disp.filter(report_period__month=month, report_period__year=year - 1)
                    qs2_disp = qs_disp.filter(report_period__month=month, report_period__year=year)
                else:
                    qs2_acm = qs_acm.filter(report__airline=airline, report__report_period__month=month,
                                            report__report_period__year=year)



                sales_data = {}
                sales_data['airline'] = airline
                if qs2:
                    # if sales_type == 'Net':
                    if is_arc_var:


                        sub_tax_cp_mf = Charges.objects.select_related('transaction').filter(type__in=['CP', 'MF'],
                                                                                             transaction__report__report_period__ped=OuterRef(
                                                                                                 'report__report_period__ped'),
                                                                                             transaction__report__airline=airline).values(

                            'transaction__report__report_period__ped').annotate(
                            dcount=Count('transaction__report__report_period__ped'),
                            total_tax_cp_mf=Sum('amount', output_field=FloatField())).values('total_tax_cp_mf')
                        t_filtr = Transaction.objects.exclude(transaction_type__startswith='SP',
                                                              agency__agency_no='6999001').select_related('report',
                                                                                                          'report__airline',
                                                                                                          'report__report_period').filter(
                            report__airline=airline,
                            report__report_period__month=month,
                            report__report_period__year=year)

                        transactions = Transaction.objects.exclude(transaction_type__startswith='SP',
                                                                   agency__agency_no='6999001').select_related(
                            'report',
                            'report__airline',
                            'report__report_period').filter(
                            report__airline=airline,
                            report__report_period__month=month,
                            report__report_period__year=year).values(
                            # 'report__report_period__ped', 'report__report_period__week', 'report__cc', 'report__ca').annotate(
                            'report__report_period__ped', 'report__report_period__week', 'card_code',
                            'report__report_period__remittance_date').annotate(
                            dcount=Count('card_code'),
                            total_fare_amount=Sum('fare_amount'),
                            total_commission=Sum('std_comm_amount'),

                            pen_value=Sum('pen'),
                            cp=Sum('pen', filter=Q(report__airline=airline, pen_type='CANCEL PEN')),

                            total_ap_acm=Sum('transaction_amount',
                                             filter=Q(report__airline=airline, transaction_type__startswith='ACM',
                                                      agency__agency_no='31768026')),

                        ).order_by(
                            'report__report_period__ped')

                        report_period_list = list(
                            transactions.values_list('report__report_period__ped', flat=True).distinct())



                        cash_fare_total = 0
                        amex_fare_total = 0
                        visa_fare_total = 0
                        mastercard_fare_total = 0
                        other_fare_total = 0

                        cash_comm_total = 0
                        amex_comm_total = 0
                        visa_comm_total = 0
                        mastercard_comm_total = 0
                        other_comm_total = 0

                        cash_tr_total = 0
                        amex_tr_total = 0
                        visa_tr_total = 0
                        mastercard_tr_total = 0
                        other_tr_total = 0

                        cash_cp_total = 0
                        amex_cp_total = 0
                        visa_cp_total = 0
                        mastercard_cp_total = 0
                        other_cp_total = 0
                        weekly_sales_value_toal = 0

                        c_total = 0
                        v_total = 0
                        m_total = 0
                        a_total = 0
                        o_total = 0

                        cash_tax_total = 0
                        amex_tax_total = 0
                        visa_tax_total = 0
                        mastercard_tax_total = 0
                        other_tax_total = 0

                        cash_yqyr_total = 0
                        amex_yqyr_total = 0
                        visa_yqyr_total = 0
                        mastercard_yqyr_total = 0
                        other_yqyr_total = 0

                        for report_period in report_period_list:
                            cash_transaction_amount = t_filtr.filter(
                                Q(transaction_type__startswith='ACM', agency__agency_no='31768026')).filter(
                                Q(card_code="") | Q(card_code=None),
                                report__report_period__ped=report_period).aggregate(
                                total_tr_value=Sum('transaction_amount')).get("total_tr_value")
                            amex_transaction_amount = t_filtr.filter(
                                Q(transaction_type__startswith='ACM', agency__agency_no='31768026')).filter(
                                card_code__startswith="3", report__report_period__ped=report_period).aggregate(
                                total_tr_value=Sum('transaction_amount')).get("total_tr_value")
                            visa_transaction_amount = t_filtr.filter(
                                Q(transaction_type__startswith='ACM', agency__agency_no='31768026')).filter(
                                card_code__startswith="4", report__report_period__ped=report_period).aggregate(
                                total_tr_value=Sum('transaction_amount')).get("total_tr_value")
                            mastercard_transaction_amount = t_filtr.filter(
                                Q(transaction_type__startswith='ACM', agency__agency_no='31768026')).filter(
                                card_code__startswith="5", report__report_period__ped=report_period).aggregate(
                                total_tr_value=Sum('transaction_amount')).get("total_tr_value")
                            other_transaction_amount = t_filtr.filter(
                                Q(transaction_type__startswith='ACM', agency__agency_no='31768026')).filter(
                                ~Q(card_code=None) & ~Q(card_code="") & ~Q(card_code="3") & ~Q(card_code="4") & ~Q(
                                    card_code="5"), report__report_period__ped=report_period).aggregate(
                                total_tr_value=Sum('transaction_amount')).get("total_tr_value")

                            if not cash_transaction_amount:
                                cash_transaction_amount = 0
                            if not amex_transaction_amount:
                                amex_transaction_amount = 0
                            if not visa_transaction_amount:
                                visa_transaction_amount = 0
                            if not mastercard_transaction_amount:
                                mastercard_transaction_amount = 0
                            if not other_transaction_amount:
                                other_transaction_amount = 0



                            cash_pen_amount = t_filtr.filter(
                                Q(card_code="") | Q(card_code=None),
                                report__report_period__ped=report_period).aggregate(
                                total_comm_value=Sum('pen')).get("total_comm_value")
                            amex_pen_amount = t_filtr.filter(
                                card_code__startswith="3", report__report_period__ped=report_period).aggregate(
                                total_comm_value=Sum('pen')).get("total_comm_value")
                            visa_pen_amount = t_filtr.filter(
                                card_code__startswith="4", report__report_period__ped=report_period).aggregate(
                                total_comm_value=Sum('pen')).get("total_comm_value")
                            mastercard_pen_amount = t_filtr.filter(
                                card_code__startswith="5", report__report_period__ped=report_period).aggregate(
                                total_comm_value=Sum('pen')).get("total_comm_value")
                            other_pen_amount = t_filtr.filter(
                                ~Q(card_code=None) & ~Q(card_code="") & ~Q(card_code__startswith="3") & ~Q(
                                    card_code__startswith="4") & ~Q(
                                    card_code__startswith="5"), report__report_period__ped=report_period).aggregate(
                                total_comm_value=Sum('pen')).get("total_comm_value")

                            amex_yq_yr_amount = Charges.objects.select_related('transaction').filter(
                                type__in=['YQ', 'YR'], transaction__report__report_period__ped=report_period,
                                transaction__report__airline=airline, transaction__card_code__startswith="3").aggregate(
                                total_comm_value=Sum('amount')).get("total_yqyr_value")
                            visa_yq_yr_amount = Charges.objects.select_related('transaction').filter(
                                type__in=['YQ', 'YR'], transaction__report__report_period__ped=report_period,
                                transaction__report__airline=airline, transaction__card_code__startswith="4").aggregate(
                                total_comm_value=Sum('amount')).get("total_yqyr_value")
                            mastercard_yq_yr_amount = Charges.objects.select_related('transaction').filter(
                                type__in=['YQ', 'YR'], transaction__report__report_period__ped=report_period,
                                transaction__report__airline=airline, transaction__card_code__startswith="5").aggregate(
                                total_comm_value=Sum('amount')).get("total_yqyr_value")
                            cash_yq_yr_amount = Charges.objects.select_related('transaction').filter(
                                Q(transaction__card_code=None) | Q(transaction__card_code=""), type__in=['YQ', 'YR'],
                                transaction__report__report_period__ped=report_period,
                                transaction__report__airline=airline).aggregate(total_comm_value=Sum('amount')).get(
                                "total_yqyr_value")
                            other_yq_yr_amount = Charges.objects.select_related('transaction').filter(
                                ~Q(transaction__card_code=None) & ~Q(transaction__card_code="") & ~Q(
                                    transaction__card_code__startswith="3") & ~Q(
                                    transaction__card_code__startswith="4") & ~Q(
                                    transaction__card_code__startswith="5"), type__in=['YQ', 'YR'],
                                transaction__report__report_period__ped=report_period).aggregate(
                                total_comm_value=Sum('amount')).get("total_yqyr_value")

                            amex_tax_amount = Taxes.objects.select_related('transaction').filter(
                                transaction__report__report_period__ped=report_period,
                                transaction__report__airline=airline, transaction__card_code__startswith="3").aggregate(
                                total_comm_value=Sum('amount')).get("total_comm_value")
                            visa_tax_amount = Taxes.objects.select_related('transaction').filter(
                                transaction__report__report_period__ped=report_period,
                                transaction__report__airline=airline, transaction__card_code__startswith="4").aggregate(
                                total_comm_value=Sum('amount')).get("total_comm_value")
                            mastercard_tax_amount = Taxes.objects.select_related('transaction').filter(
                                transaction__report__report_period__ped=report_period,
                                transaction__report__airline=airline, transaction__card_code__startswith="5").aggregate(
                                total_comm_value=Sum('amount')).get("total_comm_value")
                            cash_tax_amount = Taxes.objects.select_related('transaction').filter(
                                Q(transaction__card_code=None) | Q(transaction__card_code=""),
                                transaction__report__report_period__ped=report_period,
                                transaction__report__airline=airline).aggregate(total_comm_value=Sum('amount')).get(
                                "total_comm_value")
                            other_tax_amount = Taxes.objects.select_related('transaction').filter(
                                ~Q(transaction__card_code=None) & ~Q(transaction__card_code="") & ~Q(
                                    transaction__card_code__startswith="3") & ~Q(
                                    transaction__card_code__startswith="4") & ~Q(
                                    transaction__card_code__startswith="5"),
                                transaction__report__report_period__ped=report_period,
                                transaction__report__airline=airline).aggregate(total_comm_value=Sum('amount')).get(
                                "total_comm_value")

                            if not cash_pen_amount:
                                cash_pen_amount = 0
                            if not amex_pen_amount:
                                amex_pen_amount = 0
                            if not visa_pen_amount:
                                visa_pen_amount = 0
                            if not mastercard_pen_amount:
                                mastercard_pen_amount = 0
                            if not other_pen_amount:
                                other_pen_amount = 0

                            if not amex_yq_yr_amount:
                                amex_yq_yr_amount = 0
                            if not visa_yq_yr_amount:
                                visa_yq_yr_amount = 0
                            if not mastercard_yq_yr_amount:
                                mastercard_yq_yr_amount = 0
                            if not cash_yq_yr_amount:
                                cash_yq_yr_amount = 0
                            if not other_yq_yr_amount:
                                other_yq_yr_amount = 0

                            if not amex_tax_amount:
                                amex_tax_amount = 0
                            if not visa_tax_amount:
                                visa_tax_amount = 0
                            if not mastercard_tax_amount:
                                mastercard_tax_amount = 0
                            if not cash_tax_amount:
                                cash_tax_amount = 0
                            if not other_tax_amount:
                                other_tax_amount = 0



                            cash_cp_total += cash_pen_amount
                            amex_cp_total += amex_pen_amount
                            visa_cp_total += visa_pen_amount
                            mastercard_cp_total += mastercard_pen_amount
                            other_cp_total += other_pen_amount

                            cash_tr_total += cash_transaction_amount
                            amex_tr_total += amex_transaction_amount
                            visa_tr_total += visa_transaction_amount
                            mastercard_tr_total += mastercard_transaction_amount
                            other_tr_total += other_transaction_amount

                            cash_tax_total += cash_tax_amount
                            amex_tax_total += amex_tax_amount
                            visa_tax_total += visa_tax_amount
                            mastercard_tax_total += mastercard_tax_amount
                            other_tax_total += other_tax_amount

                            cash_yqyr_total += cash_yq_yr_amount
                            amex_yqyr_total += amex_yq_yr_amount
                            visa_yqyr_total += visa_yq_yr_amount
                            mastercard_yqyr_total += mastercard_yq_yr_amount
                            other_yqyr_total += other_yq_yr_amount


                        cash_fare_total = 0
                        amex_fare_total = 0
                        visa_fare_total = 0
                        mastercard_fare_total = 0
                        other_fare_total = 0
                        qs = Transaction.objects.select_related('agency').prefetch_related('taxes_set').filter(
                            is_sale=True).annotate(
                            yq=Sum('charges__amount', filter=Q(charges__type='YQ')),
                            yr=Sum('charges__amount', filter=Q(charges__type='YR')))
                        qs = qs.filter(report__airline=airline)
                        qs = qs.filter(report__report_period__month=month, report__report_period__year=year)

                        peds = qs.values_list('report__report_period__ped', flat=True).order_by(
                            'report__report_period__ped').distinct()

                        # total_cashhh = 0
                        for ped in peds:
                            for item in qs.filter(report__report_period__ped=ped):
                                if str(item.card_code) == "" or str(item.card_code) == "None":
                                    cash_fare_total += float(item.fare_amount) if item.fare_amount else 0
                                    cash_comm_total += float(item.std_comm_amount) if item.std_comm_amount else 0


                                elif str(item.card_code)[0] == "3":
                                    if item.fare_amount:
                                        amex_fare_total += float(item.fare_amount) if item.fare_amount else 0
                                        amex_comm_total += float(item.std_comm_amount) if item.std_comm_amount else 0


                                elif str(item.card_code)[0] == "4":
                                    if item.fare_amount:
                                        visa_fare_total += float(item.fare_amount) if item.fare_amount else 0
                                        visa_comm_total += float(item.std_comm_amount) if item.std_comm_amount else 0


                                elif str(item.card_code)[0] == "5":
                                    if item.fare_amount:
                                        mastercard_fare_total += float(item.fare_amount) if item.fare_amount else 0
                                        mastercard_comm_total += float(
                                            item.std_comm_amount) if item.std_comm_amount else 0


                                else:
                                    if item.fare_amount:
                                        other_fare_total += float(item.fare_amount) if item.fare_amount else 0
                                        other_comm_total += float(item.std_comm_amount) if item.std_comm_amount else 0

                        cash_fare_total = cash_fare_total + cash_cp_total
                        amex_fare_total = amex_fare_total + amex_cp_total
                        visa_fare_total = visa_fare_total + visa_cp_total
                        mastercard_fare_total = mastercard_fare_total + mastercard_cp_total
                        other_fare_total = other_fare_total + other_cp_total


                        cash_weekly_sales = (cash_fare_total + cash_tax_total) - cash_comm_total
                        amex_weekly_sales = (amex_fare_total + amex_tax_total) - amex_comm_total
                        visa_weekly_sales = (visa_fare_total + visa_tax_total) - visa_comm_total
                        mastercard_weekly_sales = (mastercard_fare_total + mastercard_tax_total) - mastercard_comm_total
                        other_weekly_sales = (other_fare_total + other_tax_total) - other_comm_total






                        transactions2 = t_filtr.values(
                            'report__report_period__ped', 'report__report_period__week',
                            'report__report_period__remittance_date').annotate(
                            dcount=Count('report__report_period__ped'),
                            total_fare_amount=Sum('fare_amount'),
                            total_commission=Sum('std_comm_amount'),
                            # report__cc=Sum('cc'),
                            # report__ca=Sum('ca'),
                            pen_value=Sum('pen'),
                            cp=Sum('pen', filter=Q(report__airline=airline, pen_type='CANCEL PEN')),

                            total_ap_acm=Sum('transaction_amount',
                                             filter=Q(report__airline=airline, transaction_type__startswith='ACM',
                                                      agency__agency_no='31768026')),
                            total_tax_cp_mf=Subquery(sub_tax_cp_mf, output_field=FloatField())

                        ).order_by(
                            'report__report_period__ped')

                        arc_deductions_total = 0

                        sales_data['tax_amount'] = 0.0
                        sales_data['fare_amount'] = 0.0
                        sales_data['comm_amount'] = 0.0
                        # for tt in transactions:
                        total_acm_amount = 0
                        total_cash_disbursement = 0
                        total_card_disbursement = 0
                        if transactions and transactions2:
                            visa = 0
                            visa_tax = 0
                            amex = 0
                            amex_tax = 0
                            mastercard = 0
                            mastercard_tax = 0
                            cash = 0
                            cash_tax = 0
                            mastercard_comm = 0
                            visa_comm = 0
                            amex_comm = 0
                            cash_comm = 0
                            mastercard_cpmf = 0
                            visa_cpmf = 0
                            amex_cpmf = 0
                            cash_cpmf = 0
                            mastercard_acm_fare = 0
                            visa_acm_fare = 0
                            amex_acm_fare = 0
                            cash_acm_fare = 0
                            other = 0
                            other_comm = 0
                            other_tax = 0
                            other_cpmf = 0
                            other_acm_fare = 0
                            total_ap_acm_var = 0
                            total_tax_yq_yr_total = 0

                            cash_tax_yq_yr = 0
                            amex_tax_yq_yr = 0
                            visa_tax_yq_yr = 0
                            mastercard_tax_yq_yr = 0
                            other_tax_yq_yr = 0

                            cash_total_acm_amount = 0
                            amex_total_acm_amount = 0
                            visa_total_acm_amount = 0
                            mastercard_total_acm_amount = 0
                            other_total_acm_amount = 0
                            weekly_total_sales = 0
                            # total_fare_amount = 0
                            # total_comm_amount = 0
                            total_tax_amount = 0
                            weekly_total_cash_disbursement_total = 0
                            cash_disbursement_total = 0
                            total_ap_acm = 0
                            test1 = 0
                            test2 = 0
                            test3 = 0
                            test4 = 0
                            total_weekly_sales_total = 0


                            for t_obj in transactions2:

                                total_tax_yq_yr = Charges.objects.select_related('transaction').filter(
                                    type__in=['YQ', 'YR'],
                                    transaction__report__report_period__ped=t_obj.get('report__report_period__ped'),
                                    transaction__report__airline=airline).aggregate(
                                    total_tax_yq_yr=Coalesce(Sum('amount'), V(0))).get('total_tax_yq_yr')

                                if t_obj.get('total_fare_amount'):

                                    total_fare_amount_var = t_obj.get('total_fare_amount')
                                else:
                                    total_fare_amount_var = 0.00

                                if t_obj.get('cp'):
                                    cp_var = t_obj.get('cp')
                                else:
                                    cp_var = 0.00
                                weekly_total_sales += total_fare_amount_var
                                total_fare_amount_var = total_fare_amount_var + cp_var
                                test1 += total_fare_amount_var
                                total_ap_acm_var = 0.00
                                if t_obj.get('total_ap_acm'):
                                    total_ap_acm_var = t_obj.get('total_ap_acm')
                                    total_ap_acm += total_ap_acm_var
                                test2 += total_ap_acm_var


                                total_fare_amount_var = total_fare_amount_var - total_ap_acm_var
                                test3 += total_fare_amount_var
                                total_tax_var = Taxes.objects.select_related('transaction').filter(
                                    transaction__report__report_period__ped=t_obj.get('report__report_period__ped'),
                                    transaction__report__airline=airline).aggregate(
                                    total_tax=Coalesce(Sum('amount'), V(0))).get(
                                    'total_tax')


                                if t_obj.get('total_commission'):
                                    total_commission_var = t_obj.get('total_commission')
                                else:
                                    total_commission_var = 0.00
                                weekly_sales = total_fare_amount_var + total_tax_var + total_commission_var
                                total_weekly_sales_total += weekly_sales
                                try:
                                    remittance = t.get('report__report_period__remittance_date')
                                except:
                                    pass

                                try:
                                    disb = Disbursement.objects.get(
                                        report_period__ped=t_obj.get('report__report_period__ped'),
                                        airline=airline)
                                    cash_disbursement = disb.bank7
                                    arc_deductions = disb.arc_deduction + disb.arc_fees + disb.arc_reversal + disb.arc_tot
                                    weekly_total_cash_disbursement = disb.arc_net_disb
                                except Exception as e:
                                    cash_disbursement = 0.00
                                    arc_deductions = 0.00
                                    weekly_total_cash_disbursement = 0.00



                                arc_deductions_total = arc_deductions_total + arc_deductions

                                cash_disbursement_total += cash_disbursement
                                # if weekly_total_cash_disbursement > 0:
                                weekly_total_cash_disbursement_total = weekly_total_cash_disbursement_total + weekly_total_cash_disbursement

                            total_weekly_sales = cash_weekly_sales + amex_weekly_sales + visa_weekly_sales + mastercard_weekly_sales + other_weekly_sales
                            total_card_disbursement = total_weekly_sales - cash_disbursement_total
                            cash_data = {
                                'amount': cash_fare_total,
                                'tax': cash_tax_total,
                                'comm_amount': cash_comm_total,
                                'net_sales': cash_weekly_sales
                            }
                            amex_data = {
                                'amount': amex_fare_total,
                                'tax': amex_tax_total,
                                'comm_amount': amex_comm_total,
                                'net_sales': amex_weekly_sales
                            }
                            visa_data = {
                                'amount': visa_fare_total,
                                'tax': visa_tax_total,
                                'comm_amount': visa_comm_total,
                                'net_sales': visa_weekly_sales
                            }
                            mastercard_data = {
                                'amount': mastercard_fare_total,
                                'tax': mastercard_tax_total,
                                'comm_amount': mastercard_comm_total,
                                'net_sales': mastercard_weekly_sales
                            }
                            other_data = {
                                'amount': other_fare_total,
                                'tax': other_tax_total,
                                'comm_amount': other_comm_total,
                                'net_sales': other_weekly_sales
                            }
                            card_data = {
                                'cash_data': cash_data,
                                'amex_data': amex_data,
                                'visa_data': visa_data,
                                'mastercard_data': mastercard_data,
                                'other_data': other_data
                            }
                            total_data = {
                                'total_fare': cash_fare_total + amex_fare_total + visa_fare_total + mastercard_fare_total + other_fare_total,
                                'total_tax': cash_tax_total + amex_tax_total + visa_tax_total + mastercard_tax_total + other_tax_total,
                                'total_comm': cash_comm_total + amex_comm_total + visa_comm_total + mastercard_comm_total + other_comm_total,
                                'total_net_sales': total_weekly_sales,
                                'total_acm_amount': total_ap_acm,
                                'arc_deductions_total': arc_deductions_total,
                                'total_cash_disbursement': weekly_total_cash_disbursement_total,
                                'total_card_disbursement': total_card_disbursement,
                            }
                            sales_data['card_data'] = card_data
                            sales_data['total_data'] = total_data

                            total_comm_amount = cash_comm_total + amex_comm_total + visa_comm_total + mastercard_comm_total + other_comm_total

                            agency_objs = Transaction.objects.exclude(transaction_type__startswith='SP',
                                                                      agency__agency_no='6999001').select_related(
                                'report',
                                'report__airline',
                                'report__report_period').filter(
                                report__airline=airline,
                                report__report_period__month=month,
                                report__report_period__year=year).values('agency__id', 'agency__agency_no',
                                                                         'agency__sales_owner__email',
                                                                         'agency__trade_name', 'agency__state__abrev',
                                                                         'agency__tel',
                                                                         'agency__agency_type__name').order_by().distinct().annotate(
                                dcount=Count('agency__agency_no'),
                                total_commission=Sum(
                                    'std_comm_amount')).order_by("-total_commission")[:5]

                            agency_list = []
                            for obj in agency_objs:
                                percentage = 0
                                if total_comm_amount != 0:
                                    percentage = (float(obj.get("total_commission")) / total_comm_amount) * 100
                                if obj:
                                    agency_list.append({
                                        'agency': str(obj.get("agency__agency_no")) + " - " + str(
                                            obj.get("agency__trade_name")),
                                        'amount': num_quantize(float(obj.get("total_commission"))) if obj.get(
                                            "total_commission") else 0.00,
                                        'percentage': num_quantize(abs(percentage))
                                    })

                            total_tax_var = Taxes.objects.filter(
                                transaction__report__report_period__ped__in=report_period_list,
                                transaction__report__airline=airline)
                            sales_data['agency_list'] = agency_list
                            taxes_obj = total_tax_var.values(
                                'type').order_by('type').annotate(
                                dcount=Count('type'),
                                total_amount=Sum(
                                    'amount'))


                            sales_data['tax_partials'] = taxes_obj
                    else:

                        qs2_acm_nw = qs2_acm.filter(report__airline=airline).aggregate(
                            total=Sum('fare_amount', output_field=FloatField()),
                            total_comm=Sum('std_comm_amount', output_field=FloatField()),
                            total_cp=Sum('pen', output_field=FloatField()))

                        qs2_acm_nw = qs2_acm.filter(report__airline=airline).aggregate(
                            total=Sum('transaction_amount', output_field=FloatField()),
                            total_comm=Sum('std_comm_amount', output_field=FloatField()))



                        sub_tax = Taxes.objects.select_related('transaction').filter(
                            transaction__report__report_period__ped=OuterRef('report__report_period__ped'),
                            transaction__report__airline=airline).values(
                            'transaction__report__report_period__ped').annotate(
                            dcount=Count('transaction__report__report_period__ped'),
                            total_tax=Sum('amount', output_field=FloatField())).values('total_tax')

                        sub_tax_cp_mf = Charges.objects.select_related('transaction').filter(type__in=['CP', 'MF'],
                                                                                             transaction__report__report_period__ped=OuterRef(
                                                                                                 'report__report_period__ped'),
                                                                                             transaction__report__airline=airline).values(

                            'transaction__report__report_period__ped').annotate(
                            dcount=Count('transaction__report__report_period__ped'),
                            total_tax_cp_mf=Sum('amount', output_field=FloatField())).values('total_tax_cp_mf')

                        transactions = all_transactions.filter(
                            report__airline=airline,
                            report__report_period__month=month,
                            report__report_period__year=year).values(
                            # 'report__report_period__ped', 'report__report_period__week', 'report__cc', 'report__ca').annotate(
                            'id', 'card_type', 'fare_amount', 'ticket_no', 'std_comm_amount',
                            'report__report_period__ped', 'report__report_period__week',
                            'report__report_period__remittance_date').annotate(
                            dcount=Count('report__report_period__ped'),
                            total_fare_amount=Sum('fare_amount'),
                            total_commission=Sum('std_comm_amount'),
                            report__cc=Sum('cc'),
                            report__ca=Sum('ca'),
                            pen_value=Sum('pen'),
                            cp=Sum('pen', filter=Q(report__airline=airline, pen_type='CP')),
                            total_tax=Subquery(sub_tax, output_field=FloatField()),

                            total_ap_acm=Sum('transaction_amount',
                                             filter=Q(report__airline=airline, transaction_type__startswith='ACM',
                                                      agency__agency_no='6998051')),

                            total_tax_cp_mf=Subquery(sub_tax_cp_mf, output_field=FloatField())
                        ).order_by(
                            'report__report_period__ped')
                        total_commission_amount = 0
                        for t_obj in transactions:
                            total_commission_amount += float(t_obj.get("total_commission"))
                        agency_transaction_objs = all_transactions.filter(
                            report__airline=airline,
                            report__report_period__month=month,
                            report__report_period__year=year)


                        agency_objs = agency_transaction_objs.values('agency__id', 'agency__agency_no',
                                                                     'agency__sales_owner__email',
                                                                     'agency__trade_name', 'agency__state__abrev',
                                                                     'agency__tel',
                                                                     'agency__agency_type__name').order_by().distinct().annotate(
                            dcount=Count('agency__agency_no'),
                            total_commission=Sum(
                                'std_comm_amount')).order_by("-total_commission")[:5]

                        agency_list = []
                        for obj in agency_objs:
                            percentage = 0

                            if obj:
                                if total_commission_amount != 0:
                                    percentage = (float(obj.get("total_commission")) / total_commission_amount) * 100
                                agency_list.append({
                                    'agency': str(obj.get("agency__agency_no")) + " - " + str(
                                        obj.get("agency__trade_name")),
                                    'amount': num_quantize(float(obj.get("total_commission"))) if obj.get(
                                        "total_commission") else 0.00,
                                    'percentage': num_quantize(abs(percentage))
                                })

                        sales_data['agency_list'] = agency_list
                        sales_data['tax_amount'] = 0.0
                        sales_data['fare_amount'] = 0.0
                        sales_data['comm_amount'] = 0.0
                        arc_deductions_total = 0
                        total_acm_amount = 0
                        total_cash_disbursement = 0
                        total_card_disbursement = 0
                        if transactions:
                            for t_obj in transactions:
                                if t_obj.get("total_ap_acm"):
                                    total_acm_amount += t_obj.get("total_ap_acm")

                                try:
                                    disb = Disbursement.objects.get(
                                        report_period__ped=t_obj.get('report__report_period__ped'),
                                        airline=airline)
                                    cash_disbursement = disb.bank7
                                    arc_deductions = disb.arc_deduction + disb.arc_fees + disb.arc_reversal + disb.arc_tot
                                    weekly_total_cash_disbursement = disb.arc_net_disb
                                    # pending_deduction = arc_deductions
                                except:
                                    cash_disbursement = 0.00
                                    arc_deductions = 0.00
                                    weekly_total_cash_disbursement = 0.00
                                    # pending_deduction = 0.00

                                arc_deductions_total = arc_deductions_total + arc_deductions

                                total_cash_disbursement += t_obj.get("report__ca") - t_obj.get('total_commission')
                                if t_obj.get('report__cc'):
                                    total_card_disbursement += t_obj.get('report__cc')

                            sales_data['tax_amount'] = transactions[0].get('total_tax')
                            sales_data['fare_amount'] = transactions[0].get('total_fare_amount')
                            sales_data['comm_amount'] = transactions[0].get('total_commission')
                            t = transactions[0]
                            if t.get("total_tax_cp_mf"):
                                total_tax_cp_mf = t.get("total_tax_cp_mf")
                            else:
                                total_tax_cp_mf = 0.00

                            ap_acm = all_transactions.filter(
                                report__report_period__ped=t.get('report__report_period__ped'),
                                report__airline=airline, transaction_type__startswith='ACM',
                                agency__agency_no='6998051').aggregate(
                                total=Sum('fare_amount', output_field=FloatField()),
                                total_comm=Sum('std_comm_amount', output_field=FloatField()),
                                total_cp=Sum('pen', output_field=FloatField()))
                            visa = 0
                            visa_tax = 0
                            amex = 0
                            amex_tax = 0
                            mastercard = 0
                            mastercard_tax = 0
                            cash = 0
                            cash_tax = 0
                            mastercard_comm = 0
                            visa_comm = 0
                            amex_comm = 0
                            cash_comm = 0
                            mastercard_cpmf = 0
                            visa_cpmf = 0
                            amex_cpmf = 0
                            cash_cpmf = 0
                            mastercard_acm_fare = 0
                            visa_acm_fare = 0
                            amex_acm_fare = 0
                            cash_acm_fare = 0
                            other = 0
                            other_comm = 0
                            other_tax = 0
                            other_cpmf = 0
                            other_acm_fare = 0
                            taxes_list = []
                            transaction_id_list = []

                            for card in transactions:
                                transaction_id_list.append(card.get("id"))
                                card_type = str(card.get("card_type")).replace(" ", "")
                                taxes = Taxes.objects.filter(transaction__id=card.get("id"))
                                taxes_filtered = Taxes.objects.filter(transaction__id=card.get("id"))
                                charges = Charges.objects.filter(type__in=['CP', 'MF'], transaction__id=card.get("id"))
                                acm_data = all_transactions.filter(
                                    report__report_period__ped=card.get('report__report_period__ped'),
                                    report__airline=airline, transaction_type__startswith='ACM',
                                    agency__agency_no='6998051', card_type=card.get("card_type"))
                                total_acm_fare = 0
                                total_tax = 0
                                total_charges = 0
                                for tax in taxes_filtered:
                                    taxes_list.append(tax)
                                for tax in taxes:
                                    total_tax += float(tax.amount)
                                for charge in charges:
                                    total_charges += charge.amount
                                for acm in acm_data:
                                    total_acm_fare += acm.fare_amount
                                if card_type == "Mastercard":
                                    mastercard += float(card.get("fare_amount"))
                                    mastercard_comm += float(card.get("std_comm_amount"))
                                    mastercard_tax += total_tax
                                    mastercard_cpmf += total_charges
                                    mastercard_acm_fare += total_acm_fare
                                elif card_type == "VISAInternational":
                                    visa += float(card.get("fare_amount"))
                                    visa_comm += float(card.get("std_comm_amount"))
                                    visa_tax += total_tax
                                    visa_cpmf += total_charges
                                    visa_acm_fare += total_acm_fare
                                elif card_type == "AmericanExpress":
                                    amex += float(card.get("fare_amount"))
                                    amex_comm += float(card.get("std_comm_amount"))
                                    amex_tax += total_tax
                                    amex_cpmf += total_charges
                                    amex_acm_fare += total_acm_fare
                                elif card_type != "None" and card_type not in ['AmericanExpress', 'VISAInternational',
                                                                               'Mastercard']:
                                    other += 0
                                    other_comm += float(card.get("std_comm_amount"))
                                    other_tax += total_tax
                                    other_cpmf += total_charges
                                    other_acm_fare += total_acm_fare
                                else:
                                    cash += float(card.get("fare_amount"))
                                    cash_comm += float(card.get("std_comm_amount"))
                                    cash_tax += total_tax
                                    cash_cpmf += total_charges
                                    cash_acm_fare += total_acm_fare

                            taxes_obj = Taxes.objects.filter(transaction_id__in=transaction_id_list).values(
                                'type').order_by('type').annotate(
                                dcount=Count('type'),
                                total_amount=Sum(
                                    'amount'))

                            # for obj in taxes_obj:
                            sales_data['tax_partials'] = taxes_obj
                            cash_netsales = cash - cash_comm + cash_tax
                            amex_netsales = amex - amex_comm + amex_tax
                            visa_netsales = visa - amex_comm + amex_tax
                            mastercard_netsales = mastercard - amex_comm + amex_tax
                            other_netsales = other - other_comm + other_tax
                            cash_data = {
                                'amount': cash,
                                'tax': cash_tax,
                                'comm_amount': cash_comm,
                                'net_sales': cash_netsales
                            }
                            amex_data = {
                                'amount': amex,
                                'tax': amex_tax,
                                'comm_amount': amex_comm,
                                'net_sales': amex_netsales
                            }
                            visa_data = {
                                'amount': visa,
                                'tax': visa_tax,
                                'comm_amount': visa_comm,
                                'net_sales': visa_netsales
                            }
                            mastercard_data = {
                                'amount': mastercard,
                                'tax': mastercard_tax,
                                'comm_amount': mastercard_comm,
                                'net_sales': mastercard_netsales
                            }
                            other_data = {
                                'amount': other,
                                'tax': other_tax,
                                'comm_amount': other_comm,
                                'net_sales': other_netsales
                            }
                            card_data = {
                                'cash_data': cash_data,
                                'amex_data': amex_data,
                                'visa_data': visa_data,
                                'mastercard_data': mastercard_data,
                                'other_data': other_data
                            }
                            total_data = {
                                'total_fare': cash + amex + visa + mastercard,
                                'total_tax': cash_tax + amex_tax + visa_tax + mastercard_tax,
                                'total_comm': cash_comm + amex_comm + visa_comm + mastercard_comm,
                                'total_net_sales': cash_netsales + amex_netsales + visa_netsales + mastercard_netsales,
                                'total_acm_amount': total_acm_amount,
                                'arc_deductions_total': arc_deductions_total,
                                'total_cash_disbursement': total_cash_disbursement,
                                'total_card_disbursement': total_card_disbursement,
                            }
                            sales_data['card_data'] = card_data
                            sales_data['total_data'] = total_data

                            if t.get('total_fare_amount'):
                                total_ap_acm_var = 0.00
                                if ap_acm.get('total'):
                                    total_ap_acm_var = ap_acm.get('total')
                                total_fare_amount_var = t.get('total_fare_amount') - total_ap_acm_var
                            else:
                                total_fare_amount_var = 0.00
                            if t.get('total_tax'):
                                total_tax_var = t.get('total_tax')
                            else:
                                total_tax_var = 0.00

                            total_commission_var = 0.00
                            if t.get('total_commission'):
                                total_commission_var = t.get('total_commission')

                            ap_acm_comm = 0.00
                            if ap_acm.get('total_comm'):
                                ap_acm_comm = ap_acm.get('total_comm')

                            total_commission_var = total_commission_var - ap_acm_comm
                            net_sales_weekly = total_fare_amount_var - total_commission_var + total_tax_cp_mf

                            sales_data['net_sales'] = net_sales_weekly

                    values.append(sales_data)
        sales_type = self.request.GET.get('sales_type', '')
        response = HttpResponse(content_type='application/vnd.ms-excel')

        wb = xlwt.Workbook(encoding='utf-8')
        ws = FitSheetWrapper(wb.add_sheet('All Sales for week'))
        ws.col(0).width = 250 * PT
        font_style = xlwt.XFStyle()
        font_style.font.bold = True
        # Sheet header, first row
        row_num = 0
        airline_name = ""
        if airline_id:
            airline_name = Airline.objects.get(id=airline_id).name
        if month_year and airline_id:

            file_name = "{} {}".format(airline_name, month_year) + "-weekly-sales.xls"

        else:
            file_name = "weekly-sales.xls"
            ws.row(row_num).height = 20 * 20
            row_num = row_num + 1


        row_num = row_num + 1
        ws.row(row_num).height_mismatch = True

        response['Content-Disposition'] = 'inline; filename=' + file_name


        center_normal_border = xlwt.easyxf(
            "align: wrap yes, vert centre, horiz center;font: name Arial, height 180; pattern: pattern solid, fore_color yellow;border: left thin,right thin,top thin,bottom thin")
        center_normal_border_right = xlwt.easyxf(
            "align: wrap yes, vert centre, horiz right;font: name Arial, height 180, bold True;border: left medium,right medium,top medium,bottom medium")
        center_normal_border_date = xlwt.easyxf(
            "align: wrap yes, vert centre, horiz center;font: name Arial, height 180, bold True;border: left medium,right medium,top medium,bottom medium")

        count = 1

        ws.write_merge(row_num, row_num, 0, 0, "Monthly", center_normal_border)
        # row_num += 1
        incr_num = 4
        for col_num, ele in enumerate(values, 1):

            incr_num -= 1
            ws.col(count).width = (7 * 350)
            ws.write(row_num, count, "Fare", center_normal_border)
            count = count + 1
            ws.col(count).width = (7 * 350)
            ws.write(row_num, count, "Tax", center_normal_border)
            count = count + 1
            ws.col(count).width = (7 * 350)
            ws.write(row_num, count, "Comm", center_normal_border)
            count = count + 1
            ws.col(count).width = (7 * 350)
            ws.write(row_num, count, "Net Sales", center_normal_border)
            count = count + 1
            ws.col(count).width = (7 * 350)
            ws.write(row_num, count, "ACM", center_normal_border)
            count = count + 1
            ws.col(count).width = (7 * 350)
            ws.write(row_num, count, "ARC Deduction", center_normal_border)
            count = count + 1
            ws.col(count).width = (7 * 350)
            ws.write(row_num, count, "Cash Disbursement", center_normal_border)
            count = count + 1
            ws.col(count).width = (7 * 350)
            ws.write(row_num, count, "Credit card Disbursement", center_normal_border)
            count = count + 3



        row_num = row_num + 1
        ws.write(row_num, 0, "Cash")
        ws.write(row_num + 1, 0, "AMEX")
        ws.write(row_num + 2, 0, "Visa")
        ws.write(row_num + 3, 0, "Master Card")
        ws.write(row_num + 4, 0, "Other")
        ws.write(row_num + 5, 0, "Total", center_normal_border)
        col_num = 1
        for val in values:
            card_data = val.get('card_data')
            total_data = val.get('total_data')
            cash_data = card_data.get("cash_data")
            amex_data = card_data.get("amex_data")
            visa_data = card_data.get("visa_data")
            mastercard_data = card_data.get("mastercard_data")
            other_data = card_data.get("other_data")
            # total_data = total_data.get("total_data")

            ws.write(row_num, (col_num) * 6 - 5, num_quantize(float(cash_data.get('amount'))))
            ws.write(row_num, ((col_num) * 6 - 5) + 1, num_quantize(float(cash_data.get('tax'))))
            ws.write(row_num, ((col_num) * 6 - 5) + 2, num_quantize(float(cash_data.get('comm_amount'))))
            ws.write(row_num, ((col_num) * 6 - 5) + 3, num_quantize(float(cash_data.get('net_sales'))))

            ws.write(row_num + 1, (col_num) * 6 - 5, num_quantize(float(amex_data.get('amount'))))
            ws.write(row_num + 1, ((col_num) * 6 - 5) + 1, num_quantize(float(amex_data.get('tax'))))
            ws.write(row_num + 1, ((col_num) * 6 - 5) + 2, num_quantize(float(amex_data.get('comm_amount'))))
            ws.write(row_num + 1, ((col_num) * 6 - 5) + 3, num_quantize(float(amex_data.get('net_sales'))))

            ws.write(row_num + 2, (col_num) * 6 - 5, num_quantize(float(visa_data.get('amount'))))
            ws.write(row_num + 2, ((col_num) * 6 - 5) + 1, num_quantize(float(visa_data.get('tax'))))
            ws.write(row_num + 2, ((col_num) * 6 - 5) + 2, num_quantize(float(visa_data.get('comm_amount'))))
            ws.write(row_num + 2, ((col_num) * 6 - 5) + 3, num_quantize(float(visa_data.get('net_sales'))))

            ws.write(row_num + 3, (col_num) * 6 - 5, num_quantize(float(mastercard_data.get('amount'))))
            ws.write(row_num + 3, ((col_num) * 6 - 5) + 1, num_quantize(float(mastercard_data.get('tax'))))
            ws.write(row_num + 3, ((col_num) * 6 - 5) + 2, num_quantize(float(mastercard_data.get('comm_amount'))))
            ws.write(row_num + 3, ((col_num) * 6 - 5) + 3, num_quantize(float(mastercard_data.get('net_sales'))))

            ws.write(row_num + 4, (col_num) * 6 - 5, num_quantize(float(other_data.get('amount'))))
            ws.write(row_num + 4, ((col_num) * 6 - 5) + 1, num_quantize(float(other_data.get('tax'))))
            ws.write(row_num + 4, ((col_num) * 6 - 5) + 2, num_quantize(float(other_data.get('comm_amount'))))
            ws.write(row_num + 4, ((col_num) * 6 - 5) + 3, num_quantize(float(other_data.get('net_sales'))))

            ws.write(row_num + 5, ((col_num) * 6 - 5), num_quantize(float(total_data.get('total_fare'))),
                     center_normal_border)
            ws.write(row_num + 5, ((col_num) * 6 - 5) + 1, num_quantize(float(total_data.get('total_tax'))),
                     center_normal_border)
            ws.write(row_num + 5, ((col_num) * 6 - 5) + 2, num_quantize(float(total_data.get('total_comm'))),
                     center_normal_border)
            ws.write(row_num + 5, ((col_num) * 6 - 5) + 3, num_quantize(float(total_data.get('total_net_sales'))),
                     center_normal_border)
            ws.write(row_num + 5, ((col_num) * 6 - 5) + 4, num_quantize(float(total_data.get('total_acm_amount'))),
                     center_normal_border)
            ws.write(row_num + 5, ((col_num) * 6 - 5) + 5, num_quantize(float(total_data.get('arc_deductions_total'))),
                     center_normal_border)
            ws.write(row_num + 5, ((col_num) * 6 - 5) + 6,
                     num_quantize(float(total_data.get('total_cash_disbursement'))), center_normal_border)
            ws.write(row_num + 5, ((col_num) * 6 - 5) + 7,
                     num_quantize(float(total_data.get('total_card_disbursement'))), center_normal_border)

            col_num += 1


        cl_name = 1
        row_num = row_num + 7

        large_num = 0
        ws.write_merge(row_num, row_num + 1, 0, 1,
                       "Breakup of Taxes", center_normal_border)
        ws.write_merge(row_num, row_num + 1, 4, 6,
                       "Agencywise Commission", center_normal_border)

        ws.col(0).width = (7 * 350)
        ws.write_merge(row_num + 2, row_num + 2, 0, 0, "Tax code",
                       center_normal_border)
        ws.col(1).width = (7 * 350)
        ws.write_merge(row_num + 2, row_num + 2, 1, 1,
                       "Tax amount", center_normal_border)
        ws.col(4).width = (7 * 350)
        ws.write_merge(row_num + 2, row_num + 2, 4, 4,
                       "Agency", center_normal_border)
        ws.col(5).width = (7 * 350)
        ws.write_merge(row_num + 2, row_num + 2, 5, 5,
                       "Amount", center_normal_border)
        ws.col(6).width = (7 * 350)
        ws.write_merge(row_num + 2, row_num + 2, 6, 6,
                       "%", center_normal_border)
        for val in values:
            if val.get('tax_partials'):
                num = 3
                for tax in val.get('tax_partials'):
                    try:
                        ws.write(row_num + num, 0,
                                 "TAX " + tax.get("type"))
                    except:
                        ws.write(row_num + num, 0,
                                 "TAX " + "None")

                    ws.write(row_num + num, 1,
                             num_quantize(float(tax.get("total_amount"))))
                    num += 1
        for val in values:
            if val.get("agency_list"):
                agency_num = 3
                agency_list = val.get("agency_list")

                sorted_list = agency_list[:5]
                for agency in sorted_list:
                    ws.write(row_num + agency_num, 4, agency.get("agency"))
                    ws.write(row_num + agency_num, 5, num_quantize(float(agency.get("amount"))))
                    ws.write(row_num + agency_num, 6, agency.get("percentage"))

                    agency_num += 1
        cl_name += 1


        wb.save(response)
        return response




def num_quantize(value, n_point=2):
    """
    :param value:
    :param n_point:
    :return:
    """
    from decimal import localcontext, Decimal, ROUND_HALF_UP
    with localcontext() as ctx:
        ctx.rounding = ROUND_HALF_UP
        if value:
            d_places = Decimal(10) ** -n_point
            # Round to two places
            value = Decimal(value).quantize(d_places)
        return 





###############################################################################################################
class NewSalesSummaryReport(PermissionRequiredMixin, TemplateView):
    """NewSalesSummaryReport."""

    permission_required = ('report.view_new_sales_summary',)
    template_name = 'new-sales-summary.html'

    def get_context_data(self, **kwargs):
        Total_ACM = 0
        Total_comm = 0
        Total_Penalties = 0
        Total_tax_yq_yr = 0
        Total_Balance_Payable = 0
        Total_tax_on_commission = 0
        Total_fare = 0
        Total_transaction_amount = 0
        Total_Transaction = 0
        Total_Easy_pay = 0
        Total_tax = 0
        TTL_BSP_SALES_Ex_Tax = 0
        context = super(NewSalesSummaryReport, self).get_context_data(**kwargs)
        month_year = self.request.GET.get('month_year', '')
        airline = self.request.GET.get('airline', '')

        country = self.request.session.get('country')

        is_arc_var = is_arc(self.request.session.get('country'))
        if airline and month_year:
            month = datetime.datetime.strptime(month_year, '%B %Y').month or ''

            year = datetime.datetime.strptime(month_year, '%B %Y').year or ''
        
            airline_data = Airline.objects.get(id=airline)
            airline_name = airline_data.name
            country_obj = Country.objects.get(id=country)
            country_name = country_obj.name
            currency_code = country_obj.currency
  
            iata_coordination_fee = 0.00
            gsa_commision = 0.00
            sub_tax_yq_yr = Charges.objects.select_related('transaction').filter(type__in=['YQ', 'YR'],transaction__report__report_period__ped=OuterRef('report__report_period__ped'),transaction__report__airline=airline).values('transaction__report__report_period__ped').annotate(
                dcount=Count('transaction__report__report_period__ped'),total_tax_yq_yr=Sum('amount', output_field=FloatField())).values('total_tax_yq_yr')

            sub_tax_cp_mf = Charges.objects.select_related('transaction').filter(type__in=['CP', 'MF'],
                                                                                 transaction__report__report_period__ped=OuterRef(
                                                                                     'report__report_period__ped'),
                                                                                 transaction__report__airline=airline).values(
                'transaction__report__report_period__ped').annotate(
                dcount=Count('transaction__report__report_period__ped'),
                total_tax_cp_mf=Sum('amount', output_field=FloatField())).values('total_tax_cp_mf')





            sub_tax = Taxes.objects.select_related('transaction').filter(
                transaction__report__report_period__ped=OuterRef('report__report_period__ped'),
                transaction__report__airline=airline).values(
                'transaction__report__report_period__ped').annotate(
                dcount=Count('transaction__report__report_period__ped'),
                total_tax=Sum('amount', output_field=FloatField())).values('total_tax')
        
            total_adm = AgencyDebitMemo.objects.filter(transaction__report__airline=airline,
                                                       transaction__report__report_period__ped=OuterRef(
                                                           'report__report_period__ped')).values(
                'transaction__report__report_period__ped').annotate(
                dcount=Count('transaction__report__report_period__ped'),total_adm=Sum('amount', output_field=FloatField())).values('total_adm')
 
            
            if is_arc_var:

                month_date = datetime.datetime.strptime(month_year, '%B %Y')


                commission_history = CommissionHistory.objects.filter(airline=airline, type='A',
                                                                      from_date__lte=month_date)
                if commission_history.exists():
                    iata_coordination_fee = commission_history.order_by('-from_date').first().rate
                
                commission_history = CommissionHistory.objects.filter(airline=airline, type='D',
                                                                      from_date__lte=month_date)

                if commission_history.exists():
                    gsa_commision = commission_history.order_by('-from_date').first().rate
                    

                transactions = Transaction.objects.exclude(transaction_type__startswith='SP',
                                                           agency__agency_no='6999001').select_related('report',
                                                                                                     'report__airline',
                                                                                                       'report__report_period').filter(
                    report__airline=airline,
                    report__report_period__month=month,
                    report__report_period__year=year).values(
                    'report__report_period__ped', 'report__report_period__week', 'card_type',
                    'report__report_period__remittance_date', 'report__cc', 'report__ca').annotate(
                    total_fare_amount=Sum('fare_amount'),
                    total_commission=Sum('std_comm_amount'),
                    pen_value=Sum('pen'),
                    cp=Sum('pen', filter=Q(report__airline=airline, pen_type='CANCEL PEN')),
                    total_ap_acm=Sum('transaction_amount',
                                     filter=Q(report__airline=airline, transaction_type__startswith='ACM',
                                              agency__agency_no='31768026')),

                ).order_by(
                    'report__report_period__ped')


                transactions_headers = []
                transactions_rows = []
                transactions_rows_iata = []
                transactions_rows_gsa = []
                total_due_to_airlinepros = []
                total_due_to_airlinepros_negative = []
                acms_rows = []
                acms_rows_total = []

                total_ap_acm = 0.00
                weekly_sales_total = 0.00
                total_cc = 0.00
                total_ca = 0.00

                net_sales_weekly_total = 0.00
                net_sales_weekly_total_iata = 0.00
                net_sales_weekly_total_gsa = 0.00
                pending_deduction_total = 0.00
                weekly_total_cash_disbursement_total = 0.00
                arc_deductions_total = 0.00
                cash_amount = 0
                test = 0
                add_yqyr=False



                qset= CommissionHistory.objects.filter(airline = airline).first()

                if qset:
                   add_yqyr=qset.add_yq_yr


                for t in transactions:

                    transactions_headers.append(t.get('report__report_period__ped'))

                    total_tax_yq_yr = Charges.objects.select_related('transaction').filter(type__in=['YQ', 'YR'],
                                                                                           transaction__report__report_period__ped=t.get(
                                                                                               'report__report_period__ped'),
                                                                                           transaction__report__airline=airline).aggregate(
                        total_tax_yq_yr=Coalesce(Sum('amount'), V(0))).get('total_tax_yq_yr')



                    total_tax_cp_mf = Charges.objects.select_related('transaction').filter(type__in=['CP', 'MF'],
                                                                                           transaction__report__report_period__ped=t.get(
                                                                                               'report__report_period__ped'),
                                                                                           transaction__report__airline=airline).aggregate(
                        total_tax_cp_mf=Coalesce(Sum('amount'), V(0))).get('total_tax_cp_mf')

                    penalty_charge = Charges.objects.select_related('transaction').filter(type__in=['CP'],
                                                                                           transaction__report__report_period__ped=t.get(
                                                                                               'report__report_period__ped'),
                                                                                           transaction__report__airline=airline).aggregate(
                        total_tax_cp_mf=Coalesce(Sum('amount'), V(0))).get('total_tax_cp_mf')

                    # pen_chrg_2 = Charges.objects.select_related('transaction')
                    # print("penalty charge",penalty_charge)


                    if t.get('total_fare_amount'):

                        total_fare_amount_var = t.get('total_fare_amount')
                    else:
                        total_fare_amount_var = 0.00
                        
                 

                    if t.get('cp'):
                        cp_var = t.get('cp')
                    else:
                        cp_var = 0.00

                    # total_fare_amount_var = total_fare_amount_var + cp_var
                    total_fare_amount_var = total_fare_amount_var + penalty_charge

                    total_ap_acm_var = 0.00
                    if t.get('total_ap_acm'):
                        total_ap_acm_var = t.get('total_ap_acm')

                    total_fare_amount_var = total_fare_amount_var - total_ap_acm_var
                    net_sales_weekly = total_fare_amount_var

                    total_tax_var = Taxes.objects.select_related('transaction').filter(
                        transaction__report__report_period__ped=t.get('report__report_period__ped'),
                        transaction__report__airline=airline).aggregate(total_tax=Coalesce(Sum('amount'), V(0))).get(
                        'total_tax')

                    if t.get('total_commission'):
                        total_commission_var = t.get('total_commission')*-1
                        # total_commission_var = abs(t.get('total_commission')) * -1
                    else:
                        total_commission_var = 0.00
                    weekly_sales = total_fare_amount_var + total_tax_var + total_commission_var + total_tax_yq_yr + value_check(
                        t.get('total_ap_acm'))

                    test += total_commission_var
                    remittance = t.get('report__report_period__remittance_date')


                    try:
                        disb = Disbursement.objects.get(report_period__ped=t.get('report__report_period__ped'),
                                                        airline=airline)
                        cash_disbursement = disb.bank7
                        arc_deductions = disb.arc_deduction + disb.arc_fees + disb.arc_reversal + disb.arc_tot
                        weekly_total_cash_disbursement = disb.arc_net_disb
                    except Exception as e:
                        cash_disbursement = 0.00
                        arc_deductions = 0.00
                        weekly_total_cash_disbursement = 0.00

                    try:
                        pending_deduction = Deduction.objects.filter(
                            report__report_period__ped=t.get('report__report_period__ped'), report__airline=airline,
                            pending=True).aggregate(pending_deduction=Coalesce(Sum('amount'), V(0))).get(
                            'pending_deduction')
                    except Exception as e:
                        pending_deduction = 0.00

                    arc_deductions_total = arc_deductions_total + arc_deductions
                    weekly_total_cash_disbursement_total = weekly_total_cash_disbursement_total + weekly_total_cash_disbursement
                    card_disbursement = weekly_sales - cash_disbursement

                    pending_deduction_total = pending_deduction_total + pending_deduction


                    ped_total = total_fare_amount_var
                    pen_val = t.get("pen_value")
                    if t.get("pen_value") is None:
                        pen_val = 0
                    if not total_tax_cp_mf:
                        total_tax_cp_mf = 0

                    calculated_fare_amount = total_fare_amount_var + total_tax_cp_mf - total_commission_var

                    transactions_rows.append(
                        {"fare": total_fare_amount_var, "tax": total_tax_var, "comm": total_commission_var,
                         "calculated_fare_amount": calculated_fare_amount,
                         "pen": t.get("pen_value"), "tax_yq_yr": total_tax_yq_yr, "tax_cp_mf": total_tax_cp_mf,
                         "weekly_sales_total": weekly_sales,
                         "total_ca": cash_disbursement, "total_cc": card_disbursement,
                         "total_ap_acm": t.get("total_ap_acm"), "remittance": remittance,
                         "week": t.get('report__report_period__week'), "pending_deduction": pending_deduction,
                         'net_sales_weekly': net_sales_weekly, 'ped_total': ped_total, 'arc_deductions': arc_deductions,
                         'weekly_total_cash_disbursement': weekly_total_cash_disbursement})

                    net_sales_weekly_total = net_sales_weekly_total + net_sales_weekly

                    if t.get('total_ap_acm'):
                        total_ap_acm = total_ap_acm + t.get('total_ap_acm')

                    weekly_sales_total = weekly_sales_total + weekly_sales

                    if card_disbursement:
                        total_cc = total_cc + card_disbursement
                    if cash_disbursement:
                        total_ca = total_ca + cash_disbursement

                  #ifelse for add yqyr
                    if add_yqyr :
                       if  iata_coordination_fee !=0:

                              transactions_rows_iata.append(((net_sales_weekly * iata_coordination_fee) / 100)+total_tax_yq_yr)
                              net_sales_weekly_total_iata = net_sales_weekly_total_iata + ((
                        (net_sales_weekly * iata_coordination_fee) / 100)+total_tax_yq_yr)
                       else:
                           transactions_rows_iata.append((net_sales_weekly * iata_coordination_fee) / 100)

                       if gsa_commision !=0:

                        transactions_rows_gsa.append(((net_sales_weekly * gsa_commision) / 100) + total_tax_yq_yr )
 
                       else:
                           transactions_rows_gsa.append((net_sales_weekly * gsa_commision) / 100)


                       net_sales_weekly_total_gsa = net_sales_weekly_total_gsa + (
                        (net_sales_weekly * gsa_commision) / 100)

                       total_due_to_airlinepros.append(
                            ((net_sales_weekly * iata_coordination_fee) / 100) + ((net_sales_weekly * gsa_commision) / 100))
                       total_due_to_airlinepros_negative.append(((((net_sales_weekly * iata_coordination_fee) / 100) + (
                            (net_sales_weekly * gsa_commision) / 100))) * -1)

    

                    else:
                        # ifelse iata
                        transactions_rows_iata.append((net_sales_weekly * iata_coordination_fee) / 100)
                        transactions_rows_gsa.append((net_sales_weekly * gsa_commision) / 100)
 


                        net_sales_weekly_total_iata = net_sales_weekly_total_iata + (
                            (net_sales_weekly * iata_coordination_fee) / 100)
                        net_sales_weekly_total_gsa = net_sales_weekly_total_gsa + (
                            (net_sales_weekly * gsa_commision) / 100)

                        total_due_to_airlinepros.append(
                            ((net_sales_weekly * iata_coordination_fee) / 100) + ((net_sales_weekly * gsa_commision) / 100))
                        total_due_to_airlinepros_negative.append(((((net_sales_weekly * iata_coordination_fee) / 100) + (
                            (net_sales_weekly * gsa_commision) / 100))) * -1)


            else:
                # Transaction.objects.filter(report__airline__name = "TAAG Angola",report__country__name = "Canada",report__report_period__week= 1, report__report_period__year= 2023,report__report_period__month= 2).delete()

                month_date = datetime.datetime.strptime(month_year, '%B %Y')

                commission_history = CommissionHistory.objects.filter(airline=airline, type='I',
                                                                  from_date__lte=month_date)





                if commission_history.exists():
                    iata_coordination_fee = commission_history.order_by('-from_date').first().rate
                    
      
        
                commission_history = CommissionHistory.objects.filter(airline=airline, type='G',
                                                                      from_date__lte=month_date)
            
                if commission_history.exists():

                    gsa_commision = commission_history.order_by('-from_date').first().rate


                transactions = Transaction.objects.exclude(transaction_type__startswith='SP',
                                                           agency__agency_no='6999001').select_related('report',
                                                                                                   'report__airline',
                                                                                                       'report__report_period').filter(
                    report__airline=airline,
                    report__report_period__month=month,
                    report__report_period__year=year).values(
                    'report__report_period__ped', 'report__report_period__week',
                    'report__report_period__remittance_date').annotate(
                    dcount=Count('report__report_period__ped'),
                    total_transaction_amount=Sum('transaction_amount'),
                    total_fare_amount=Sum('fare_amount'),
                    total_commission=Sum('std_comm_amount'),
                    tax_on_commission=Sum('tax_on_comm'),
                    report__cc=Sum('cc'),
                    report__ca=Sum('ca'),
                    ep=Sum('ep'),
                    Balance_Payable=Sum('balance'),
                    pen_value=Sum('pen'),
                    cp=Sum('pen', filter=Q(report__airline=airline, pen_type='CP')),
                    total_tax=Subquery(sub_tax, output_field=FloatField()),
                    total_adm=Subquery(total_adm, output_field=FloatField()),
                    total_tax_yq_yr=Subquery(sub_tax_yq_yr,
                                             output_field=FloatField()),
                    total_tax_cp_mf=Subquery(sub_tax_cp_mf, output_field=FloatField())
                ).order_by(
                    'report__report_period__ped')
                
               
                # transactions=transactions.exclude(transaction_type__startswith='ADM')
                
                transactions_headers = []
                transactions_rows = []
                transactions_rows_iata = []
                transactions_rows_gsa = []
                total_due_to_airlinepros = []
                total_due_to_airlinepros_negative = []
                acms_rows = []
                acms_rows_total = []

                total_ap_acm = 0.00
                weekly_sales_total = 0.00
                total_cc = 0.00
                total_ca = 0.00

                net_sales_weekly_total = 0.00
                net_sales_weekly_total_iata = 0.00
                net_sales_weekly_total_gsa = 0.00
                pending_deduction_total = 0.00
                weekly_total_cash_disbursement_total = 0.00
                arc_deductions_total = 0.00
                total_ap_acm_var = 0.00


                if not transactions:

                    for c_t in range(0, 4):
            
                        t = {}
                        ap_acm = {}
                        month_day = self.get_all_sundays(month, year)[c_t]
                        transactions_headers.append(month_day)
       
                        total_tax_yq_yr = 0.00
                        total_tax_cp_mf = 0.00
                        total_fare_amount_var = 0.00
                       
                        total_commission_var = 0.00
                        total_tax_var = 0.00
                        ap_acm_comm = 0.00
                        total_commission_var = 0.00
                        cp_var = 0.00
                        ap_acm_cp = 0.00
                        cp_var = 0.00


                        remittance = month_day

                        cash_disbursement = 0.00

                        pending_deduction_total = 0.00

                        net_sales_weekly = 0.00

                        weekly_sales = 0.00
                        pen_val = 0
                        if t.get("pen_value"):
                            pen_val = float(t.get("pen_value"))
                        calculated_fare_amount = "{:.2f}".format(
                            total_fare_amount_var + total_tax_cp_mf - total_commission_var)

                        transactions_rows.append(
                            {"fare": total_fare_amount_var,"tax": t.get('total_tax'),
                             "calculated_fare_amount": calculated_fare_amount,
                             "pen": t.get("pen_value"), "comm": (abs(total_commission_var) * -1),
                             "tax_yq_yr": total_tax_yq_yr, "tax_cp_mf": total_tax_cp_mf,
                             "weekly_sales_total": weekly_sales,
                             "total_ca": cash_disbursement, "total_cc": t.get("report__cc"),
                             "total_ap_acm": total_ap_acm_var, "remittance": remittance,
                             "week": c_t + 1, "pending_deduction": 0.00,
                             'net_sales_weekly': net_sales_weekly, 'ped_total': total_fare_amount_var})
                     
                        net_sales_weekly_total = net_sales_weekly_total + net_sales_weekly

                        total_ap_acm = 0.00

                        weekly_sales_total = 0.00

                        total_cc = 0.00
                        total_ca = 0.00

                        transactions_rows_iata.append((net_sales_weekly * iata_coordination_fee) / 100)
                        transactions_rows_gsa.append((net_sales_weekly * gsa_commision) / 100)
                     
                        net_sales_weekly_total_iata = net_sales_weekly_total_iata + (
                            (net_sales_weekly * iata_coordination_fee) / 100)
                        net_sales_weekly_total_gsa = net_sales_weekly_total_gsa + (
                            (net_sales_weekly * gsa_commision) / 100)

                        total_due_to_airlinepros.append(
                            ((net_sales_weekly * iata_coordination_fee) / 100) + (
                                (net_sales_weekly * gsa_commision) / 100))
                        total_due_to_airlinepros_negative.append(
                            ((((net_sales_weekly * iata_coordination_fee) / 100) + (
                                (net_sales_weekly * gsa_commision) / 100))) * -1)
                     
                c_t = 0
                Total_transaction_amount = 0
                Total_fare = 0
                Total_tax = 0
                Total_tax_yq_yr = 0
                Total_Penalties = 0
                Total_Easy_pay = 0
                Total_comm = 0
                Total_tax_on_commission = 0
                Total_Balance_Payable = 0
                Total_amd = 0
                Total_cc = 0
                Total_ca = 0
                ACSA_Fee = 0
                
                total_Transaction = Transaction.objects.filter(report__airline = airline,report__report_period__month = month,report__report_period__year= year,report__country = country ).exclude(transaction_type="+RTDN")
               
                li = []
                for i in total_Transaction:
                    li.append(i.ticket_no)
                Total_Transaction = len(set(li))
                
                
                for t in transactions:

                    transactions_headers.append(t.get('report__report_period__ped'))
         
                    


                    if t.get("total_tax_yq_yr"):
                        total_tax_yq_yr = t.get("total_tax_yq_yr")
          

                    else:
                        total_tax_yq_yr = 0.00

                    if t.get("total_tax_cp_mf"):
                        total_tax_cp_mf = t.get("total_tax_cp_mf")
          
                    else:
                        total_tax_cp_mf = 0.00
                        
                    
                       
    
                    ap_acm = Transaction.objects.filter(report__report_period__ped=t.get('report__report_period__ped'),
                                                        report__airline=airline, transaction_type__startswith='ACM',
                                                        agency__agency_no='6998051').aggregate(
                        
                        total=Sum('fare_amount', output_field=FloatField()),
                        total_comm=Sum('std_comm_amount', output_field=FloatField()),
                        total_cp=Sum('pen', output_field=FloatField()))

                    total_cp = 0
                    for ac in Transaction.objects.filter(report__report_period__ped=t.get('report__report_period__ped'),
                                                         report__airline=airline, transaction_type__startswith='ACM',
                                                         agency__agency_no='6998051'):
                        for chrge in ac.charges_set.filter(type__in=['CP', 'MF']):
                            total_cp += int(chrge.amount)

                    if t.get('total_fare_amount'):
                        total_ap_acm_var = 0
                        if ap_acm.get('total'):
                            total_ap_acm_var = ap_acm.get('total')
                        total_fare_amount_var = t.get('total_fare_amount') + total_cp
                    else:
                        total_fare_amount_var = 0.00

                    if t.get('total_adm'):
                        total_adm = t.get('total_adm')
                    else:
                        total_adm = 0.00
                        
                    if t.get('total_transaction_amount'):
                        total_transaction_amount_var = t.get('total_transaction_amount')
                    else:
                        total_transaction_amount_var = 0.00
                    
                    if t.get('Balance_Payable'):
                        balance_Payable_var = t.get('Balance_Payable')
                    else:
                        balance_Payable_var = 0.00
                    
                    if t.get('tax_on_commission'):
                        tax_on_commission_var = t.get('tax_on_commission')
                    else:
                        tax_on_commission_var = 0.00
                      
                    
                    if t.get('ep'):
                        Easy_pay_var = t.get('ep')
                    else:
                        Easy_pay_var = 0.00 
                    

                    if t.get('total_tax'):
                        total_tax_var = t.get('total_tax')
                    else:
                        total_tax_var = 0.00
                    total_commission_var = 0.00
                    if t.get('total_commission'):
                        total_commission_var = t.get('total_commission')

                    ap_acm_comm = 0.00
                    if ap_acm.get('total_comm'):
                        ap_acm_comm = ap_acm.get('total_comm')

                    total_commission_var = total_commission_var - ap_acm_comm

                    cp_var = 0.00
                    if t.get('cp'):
                        cp_var = t.get('cp')


                    ap_acm_cp = 0.00

                    if total_cp:
                        ap_acm_cp = total_cp
                    cp_var = cp_var - ap_acm_cp



                    remittance = t.get('report__report_period__remittance_date')

                    cash_disbursement = t.get("report__ca") - t.get('total_commission')

                    pending_deduction = value_check(t.get("report__ca")) + value_check(t.get('total_adm')) - (
                        value_check(t.get('total_commission'))) - value_check(t.get('total_ap_acm'))

                    pending_deduction_total = pending_deduction_total + pending_deduction

                    try:
                        country_obj = Country.objects.get(id=country)
                        if country_obj.name.lower() in ['canada', 'united states']:
                            net_sales_weekly = total_fare_amount_var - total_commission_var + total_tax_cp_mf
                        else:
                            net_sales_weekly = t.get("total_fare_amount")
                    except:
                        net_sales_weekly = total_fare_amount_var - total_commission_var + total_tax_cp_mf
                    weekly_sales = net_sales_weekly + total_tax_var + total_tax_yq_yr
  

                    pen_val = t.get("pen_value")
                  
                    if t.get("pen_value") is None:
                        pen_val = 0
                    calculated_fare_amount = total_fare_amount_var + total_tax_cp_mf - total_commission_var

                    Total_transaction_amount = Total_transaction_amount + total_transaction_amount_var
                    Total_fare = Total_fare + total_fare_amount_var
    
                    if not t.get('total_tax') :
                        Total_tax = Total_tax + 0
                    else:
                        Total_tax = Total_tax + t.get('total_tax')
                    Total_tax_yq_yr = Total_tax_yq_yr + total_tax_yq_yr
                    Total_Penalties = Total_Penalties + total_tax_cp_mf
                    Total_Easy_pay = Total_Easy_pay + Easy_pay_var
                    Total_comm = Total_comm + (abs(total_commission_var))
                    Total_tax_on_commission = Total_tax_on_commission + tax_on_commission_var
                    Total_Balance_Payable = Total_Balance_Payable + balance_Payable_var
                    Total_ca = Total_ca + (balance_Payable_var+(abs(total_commission_var)))
                    Total_cc = Total_cc + t.get("report__cc")
                    Total_amd = Total_amd + total_adm
                    TTL_BSP_SALES_Ex_Tax = Total_cc + Total_ca + Total_Easy_pay - Total_tax
                    
                    
                    transactions_rows.append(
                        {"fare": total_fare_amount_var,"Transaction_amount":total_transaction_amount_var, "tax": t.get('total_tax'),
                         "calculated_fare_amount": calculated_fare_amount,"Balance_Payable":balance_Payable_var,
                         "comm": (abs(total_commission_var)), "pen": t.get("pen_value"),"tax_on_commission": tax_on_commission_var,
                         "tax_yq_yr": total_tax_yq_yr, "tax_cp_mf": total_tax_cp_mf, "weekly_sales_total": weekly_sales,
                         "total_ca": (balance_Payable_var+(abs(total_commission_var))), "total_cc": t.get("report__cc"), "Easy_pay": Easy_pay_var,
                         "total_ap_acm": total_ap_acm_var, "remittance": remittance,
                         "week": t.get('report__report_period__week'), "pending_deduction": pending_deduction,
                         'net_sales_weekly': net_sales_weekly, 'ped_total': total_fare_amount_var,"TTL_BSP_SALES_Ex_Tax": TTL_BSP_SALES_Ex_Tax,
                         "ACSA_Fee": ACSA_Fee})
                
                    net_sales_weekly_total = net_sales_weekly_total + net_sales_weekly

                    if total_ap_acm_var:
                        total_ap_acm = total_ap_acm + total_ap_acm_var

                    weekly_sales_total = weekly_sales_total + weekly_sales

                    if t.get('report__cc'):
                        total_cc = total_cc + t.get('report__cc')
                    if cash_disbursement:
                        total_ca = total_ca + cash_disbursement


                    transactions_rows_iata.append((net_sales_weekly * iata_coordination_fee) / 100)
                    transactions_rows_gsa.append((net_sales_weekly * gsa_commision) / 100)

                    net_sales_weekly_total_iata = net_sales_weekly_total_iata + (
                        (net_sales_weekly * iata_coordination_fee) / 100)
                    net_sales_weekly_total_gsa = net_sales_weekly_total_gsa + (
                        (net_sales_weekly * gsa_commision) / 100)

                    total_due_to_airlinepros.append(
                        ((net_sales_weekly * iata_coordination_fee) / 100) + ((net_sales_weekly * gsa_commision) / 100))
                    total_due_to_airlinepros_negative.append(((((net_sales_weekly * iata_coordination_fee) / 100) + (
                        (net_sales_weekly * gsa_commision) / 100))) * -1)
                
                Total_ACM = sum(total_due_to_airlinepros_negative)
                
                    
 
                
               
                 
                if transactions and len(transactions) < 4:
                    m_c = [1, 2, 3, 4]
                    for i in transactions:
                        try:
                            m_c.remove(i.get('report__report_period__week'))
                        except:
                            pass
                    c_t = m_c[0]

                    t = {}
                    ap_acm = {}

                    month_day = self.get_all_sundays(month, year)[c_t - 1]
                    transactions_headers.insert(c_t - 1, month_day)

                    total_tax_yq_yr = 0.00
                    total_fare_amount_var = 0.00
                    total_commission_var = 0.00
                    total_tax_var = 0.00
                    ap_acm_comm = 0.00
                    total_commission_var = 0.00
                    cp_var = 0.00
                    ap_acm_cp = 0.00
                    cp_var = 0.00
                    total_adm = 0.00
                    total_transaction_amount_var = 0.00
                    Easy_pay_var = 0.00
                    tax_on_commission_var = 0.00
                    balance_Payable_var = 0.00
                    
                    remittance = month_day

                    cash_disbursement = 0.00

                    pending_deduction_total = 0.00

                    net_sales_weekly = 0.00

                    weekly_sales = 0.00

                    transactions_rows.insert(c_t - 1,
                                             {"fare": total_fare_amount_var, "tax": t.get('total_tax'),
                                              "comm": (abs(total_commission_var) * -1),
                                              "tax_yq_yr": total_tax_yq_yr, "weekly_sales_total": weekly_sales,
                                              "total_ca": cash_disbursement, "total_cc": t.get("report__cc"),
                                              "total_ap_acm": total_ap_acm_var, "remittance": remittance,
                                              "week": c_t, "pending_deduction": 0.00,
                                              'net_sales_weekly': net_sales_weekly, 'ped_total': total_fare_amount_var})


                    net_sales_weekly_total = net_sales_weekly_total + net_sales_weekly
               
                    total_ap_acm = 0.00

                    weekly_sales_total = 0.00

                    total_cc = 0.00
                    total_ca = 0.00

                    transactions_rows_iata.insert(c_t - 1, (net_sales_weekly * iata_coordination_fee) / 100)
                    transactions_rows_gsa.insert(c_t - 1, (net_sales_weekly * gsa_commision) / 100)

                    net_sales_weekly_total_iata = net_sales_weekly_total_iata + (
                        (net_sales_weekly * iata_coordination_fee) / 100)
                    net_sales_weekly_total_gsa = net_sales_weekly_total_gsa + (
                        (net_sales_weekly * gsa_commision) / 100)

                    total_due_to_airlinepros.insert(c_t - 1,
                                                    ((net_sales_weekly * iata_coordination_fee) / 100) + (
                                                        (net_sales_weekly * gsa_commision) / 100))
                    total_due_to_airlinepros_negative.insert(c_t - 1,
                                                             ((((net_sales_weekly * iata_coordination_fee) / 100) + (
                                                                 (net_sales_weekly * gsa_commision) / 100))) * -1)
            months_weekend = ReportPeriod.objects.filter(year=year, month=month, country=country).values_list('ped',
                                                                                                              flat=True).order_by(
                'ped')
            file_uploaded_date = ReportFile.objects.select_related('airline', 'report_period').filter(airline=airline,
                                                                                                      report_period__month=month,
                                                                                                      report_period__year=year,
                                                                                                      country=country).values_list(
                'report_period__ped', flat=True).order_by('report_period__ped')

            try:
                first_ped = months_weekend[0]
                last_ped = months_weekend[len(months_weekend) - 1]
                start = first_ped - datetime.timedelta(days=5)
                end = last_ped + datetime.timedelta(days=1)
                days = [start + datetime.timedelta(day) for day in range((end - start).days + 1)]

                details_daily = DailyCreditCardFile.objects.filter(airline=airline, date__gte=start,
                                                                   date__lte=end).values_list('date', flat=True)

                missing_dates = set(months_weekend) - set(file_uploaded_date)

                missing_dates_daily = set(days) - set(details_daily)

                if is_arc_var:
                    not_uploadedfiles = []
                    for dis in Disbursement.objects.filter(airline=airline, filedate__gte=start,
                                                           filedate__lte=end):
                        if not dis.file1 or not dis.file2:
                            not_uploadedfiles.append(dis)
            except IndexError as e:
                not_uploadedfiles = []
                missing_dates = {}
                missing_dates_daily = {}
                messages.add_message(self.request, messages.WARNING,
                                     'Calendar file not uploaded for this period.')

            credit_file_dates = CarrierDeductions.objects.filter(airline=airline, filedate__month=month,
                                                                 filedate__year=year).values_list('filedate', flat=True)
            disbursement_files = Disbursement.objects.filter(airline=airline, filedate__month=month,
                                                             filedate__year=year).values_list('filedate', flat=True)

            if is_arc_var:
                missing_dates_credit = []
                if months_weekend.exists():
                    missing_dates_credit_val = months_weekend[0]
                    if missing_dates_credit_val not in credit_file_dates:
                        missing_dates_credit = [missing_dates_credit_val]
                missing_dates_disb = set(months_weekend) - set(disbursement_files)
                missing_dates_daily = []

            else:
                not_uploadedfiles = []
                missing_dates_credit = []
                missing_dates_disb = []

            TTL_BSP_SALES_Ex_Tax = Total_cc + Total_ca + Total_Easy_pay - Total_tax
            if country == '12' or  country == '13':    
                if airline == '126' or airline == '132':
                    ACSA_Fee = TTL_BSP_SALES_Ex_Tax * (3/100)
                else:
                    ACSA_Fee = TTL_BSP_SALES_Ex_Tax * (iata_coordination_fee/100)
            else:
                ACSA_Fee = TTL_BSP_SALES_Ex_Tax * (iata_coordination_fee/100)
            
           
            month_year = self.request.GET.get('month_year', '')
            airline = self.request.GET.get('airline', '')
            country = self.request.session.get('country')
            self.is_arc_var = is_arc(self.request.session.get('country'))

            # adms = []
            
            # if airline or month_year:

            #     if not self.is_arc_var:
            #         trans_type = 'TKTT'
            #     else:
            #         trans_type = 'TKT'

            #     qs = Transaction.objects.filter(transaction_type=trans_type)
            #     adm_objs = AgencyDebitMemo.objects.select_related('transaction').filter(
            #         transaction__transaction_type=trans_type)
                
            #     if airline:
            #         qs = qs.filter(report__airline=airline)
            #         #

            #         adm_objs = adm_objs.filter(transaction__report__airline=airline)
            #     if month_year:
            #         month = datetime.datetime.strptime(month_year, '%B %Y').month or ''
            #         year = datetime.datetime.strptime(month_year, '%B %Y').year or ''
            #         month_date = datetime.datetime.strptime(month_year, '%B %Y')
            #         if month and year:
            #             qs = qs .filter(report__report_period__month=month,
            #                         report__report_period__year=year)
            #             #

            #             adm_objs = adm_objs.filter(transaction__report__report_period__month=month,
            #                                 transaction__report__report_period__year=year)
            #             #

            #             if qs.count() + adm_objs.count() > 10000:
            #                 # 10000:
            #                 base_url = self.request.scheme + '://' + self.request.get_host()
            #                 excel_adm_report(country, month_year, airline, self.request.user.email, base_url)
            #                 messages.add_message(self.request, messages.WARNING,
            #                                     'Excel file generation is taking more time than expected. You will receive an email with the link to download the file once it is done.')
            #                 return HttpResponseRedirect('/reports/adm/?' + self.request.META.get('QUERY_STRING'))

            #             adms_list = adm_objs.values_list('transaction__ticket_no', flat=True)
            #             # ---------------------------------------------------------------------------


            #             allowed_commission_rate = 0.00
            #             Total_amd = 0
            #             for obj in qs:
            #                 if obj.ticket_no not in adms_list  :
            #                     #  and obj.agency.agency_no not in exclude_list
            #                     allowed_commission_rate = 0.00

            #                     commission_history = CommissionHistory.objects.filter(airline=airline, type='M',
            #                                                                         from_date__lte=obj.report.report_period.ped)
            #                     if commission_history.exists():
            #                         allowed_commission_rate = commission_history.order_by('-from_date').first().rate

            #                     try:
            #                         taken_commission_rate = obj.std_comm_rate
            #                         commission_rate_diff = allowed_commission_rate - float(taken_commission_rate)
            #                         cobl_amount = obj.fare_amount
            #                         if cobl_amount:
            #                             if commission_rate_diff < 0:
            #                                 # ADM
            #                                 adm_amount = (abs(commission_rate_diff) * cobl_amount) / 100
            #                                 Total_amd = Total_amd + adm_amount
            #                                 comment = "Commission deducted " + str(
            #                                     taken_commission_rate) + "%. Carrier authorized " + str(
            #                                     allowed_commission_rate) + "%"
            #                                 allowed_commission_amount = (cobl_amount * allowed_commission_rate) / 100
            #                                 adms.append({
            #                                     'agency_no': obj.agency.agency_no,
            #                                     'trade_name': obj.agency.trade_name,
            #                                     'ticket_no': obj.ticket_no,
            #                                     'issue_date': obj.issue_date,
            #                                     'fare_amount': obj.fare_amount,
            #                                     'std_comm_amount': obj.std_comm_amount,
            #                                     'std_comm_rate': obj.std_comm_rate,
            #                                     'allowed_commission_amount': allowed_commission_amount,
            #                                     'amount': adm_amount,
            #                                     'comment': comment

            #                                 })
            #                     except Exception as e:
            #                         pass
                        
                                                
                        
            
            
            context['currency_code'] = currency_code
            context['transactions_headers'] = transactions_headers
            context['transactions_rows'] = transactions_rows
            context['missing_dates'] = sorted(missing_dates)
            context['missing_dates_credit'] = sorted(missing_dates_credit)
            context['missing_dates_disb'] = sorted(missing_dates_disb)
            context['not_uploadedfiles'] = not_uploadedfiles
            context['missing_dates_daily'] = sorted(missing_dates_daily)
            context['missing_dates_count'] = len(context['missing_dates']) + len(context['missing_dates_disb']) + len(
                context['missing_dates_credit']) + len(context['missing_dates_daily']) + len(context['not_uploadedfiles'])
            context['transactions_rows_iata'] = transactions_rows_iata
            context['transactions_rows_gsa'] = transactions_rows_gsa
            context['total_due_to_airlinepros'] = total_due_to_airlinepros
            context['weekly_total_cash_disbursement_total'] = weekly_total_cash_disbursement_total
            context['arc_deductions_total'] = arc_deductions_total
            context['total_due_to_airlinepros_negative'] = total_due_to_airlinepros_negative
            context['net_sales_weekly_total'] = round(net_sales_weekly_total, 2)
            
            context['Total_cc'] = Total_cc
            context['Total_ca'] = Total_ca
            context['Total_amd'] = Total_amd
            context['Total_Transaction'] = Total_Transaction
            context['ACSA_Fee'] = ACSA_Fee
            context['Total_ACM'] = Total_ACM
            context['Total_transaction_amount'] = Total_transaction_amount
            context['Total_fare'] = Total_fare
            context['Total_tax'] = Total_tax
            context['Total_tax_yq_yr'] = Total_tax_yq_yr
            context['Total_Penalties'] = Total_Penalties
            context['Total_Easy_pay'] = Total_Easy_pay
            context['Total_comm'] = Total_comm
            context['Total_tax_on_commission'] = Total_tax_on_commission
            context['Total_Balance_Payable'] = Total_Balance_Payable
            context['TTL_BSP_SALES_Ex_Tax'] = TTL_BSP_SALES_Ex_Tax
            context['total_ap_acm'] = total_ap_acm
            context['weekly_sales_total'] = weekly_sales_total
            context['total_cc'] = total_cc
            context['total_ca'] = total_ca
            context['net_sales_weekly_total_iata'] = net_sales_weekly_total_iata
            context['net_sales_weekly_total_gsa'] = net_sales_weekly_total_gsa
            context['sum_total_due_to_airlinepros'] = net_sales_weekly_total_gsa + net_sales_weekly_total_iata
            context['pending_deduction_total'] = pending_deduction_total
            context['iata_coordination_fee'] = iata_coordination_fee

            context['gsa_commision'] = gsa_commision
        context['month_year'] = month_year
        context['selected_airline'] = airline
        context['airlines'] = Airline.objects.filter(country=self.request.session.get('country'))
        context['country'] = country
        return context
        

    def get_all_sundays(self, month, year):
        import calendar

        sundays = []
        cal = calendar.Calendar()

        for day in cal.itermonthdates(year, month):
            if day.weekday() == 6 and day.month == month:
                sundays.append(day)
        return sundays
########################################################################################################################################



class NewGetSalesSummaryReport(PermissionRequiredMixin, View):
    """Filtered SalesSummary Report download as CSV."""

    permission_required = ('report.download_new_sales_summary',)

    def get(self, request, *args, **kwargs):
        month_year = self.request.GET.get('month_year', '')
        airline = self.request.GET.get('airline', '')
        airline_name = ''
        is_arc_var = is_arc(self.request.session.get('country'))
        response = HttpResponse(content_type='application/vnd.ms-excel')
        country = self.request.session.get('country')
        if airline and month_year:
            
            month = datetime.datetime.strptime(month_year, '%B %Y').month or ''
            year = datetime.datetime.strptime(month_year, '%B %Y').year or ''
            airline_data = Airline.objects.get(id=airline)
            airline_name = airline_data.name
            country_obj = Country.objects.get(id=country)
            country_name = country_obj.name
            currency_code = country_obj.currency
            
            response[
                'Content-Disposition'] = 'inline; filename=' + airline_data.abrev + ' BSP-Sales Summary ' + month_year +"- {} ({})".format(country_name,currency_code)+ '.xls'

            iata_coordination_fee = 0.00
            gsa_commision = 0.00

            sub_tax_yq_yr = Charges.objects.select_related('transaction').filter(type__in=['YQ', 'YR'],
                                                                                 transaction__report__report_period__ped=OuterRef(
                                                                                     'report__report_period__ped'),
                                                                                 transaction__report__airline=airline).values(
                'transaction__report__report_period__ped').annotate(
                dcount=Count('transaction__report__report_period__ped'),
                total_tax_yq_yr=Sum('amount', output_field=FloatField())).values('total_tax_yq_yr')

            sub_tax_cp_mf = Charges.objects.select_related('transaction').filter(type__in=['CP', 'MF'],
                                                                                 transaction__report__report_period__ped=OuterRef(
                                                                                     'report__report_period__ped'),
                                                                                 transaction__report__airline=airline).values(

                'transaction__report__report_period__ped').annotate(
                dcount=Count('transaction__report__report_period__ped'),
                total_tax_cp_mf=Sum('amount', output_field=FloatField())).values('total_tax_cp_mf')
            sub_tax = Taxes.objects.select_related('transaction').filter(
                transaction__report__report_period__ped=OuterRef('report__report_period__ped'),
                transaction__report__airline=airline).values(
                'transaction__report__report_period__ped').annotate(
                dcount=Count('transaction__report__report_period__ped'),
                total_tax=Sum('amount', output_field=FloatField())).values('total_tax')

            total_adm = AgencyDebitMemo.objects.filter(transaction__report__airline=airline,
                                                       transaction__report__report_period__ped=OuterRef(
                                                           'report__report_period__ped')).values(
                'transaction__report__report_period__ped').annotate(
                dcount=Count('transaction__report__report_period__ped'),
                total_adm=Sum('amount', output_field=FloatField())).values('total_adm')
            if is_arc_var:
                
                month_date = datetime.datetime.strptime(month_year, '%B %Y')

                commission_history = CommissionHistory.objects.filter(airline=airline, type='A',
                                                                      from_date__lte=month_date)
                if commission_history.exists():
                    iata_coordination_fee = commission_history.order_by('-from_date').first().rate

                commission_history = CommissionHistory.objects.filter(airline=airline, type='D',
                                                                      from_date__lte=month_date)
                if commission_history.exists():
                    gsa_commision = commission_history.order_by('-from_date').first().rate

                transactions = Transaction.objects.exclude(transaction_type__startswith='SP',
                                                           agency__agency_no='6999001').select_related('report',
                                                                                                       'report__airline',
                                                                                                       'report__report_period').filter(
                    report__airline=airline,
                    report__report_period__month=month,
                    report__report_period__year=year).values(
                    'report__report_period__ped', 'report__report_period__week',
                    'report__report_period__remittance_date', 'report__cc', 'report__ca').annotate(
                    total_fare_amount=Sum('fare_amount'),
                    total_commission=Sum('std_comm_amount'),
                    cp=Sum('pen', filter=Q(report__airline=airline, pen_type='CANCEL PEN')),

                    total_ap_acm=Sum('transaction_amount',
                                     filter=Q(report__airline=airline, transaction_type__startswith='ACM',
                                              agency__agency_no='31768026')),

                ).order_by(
                    'report__report_period__ped')

                transactions_headers = []
                transactions_rows = []
                transactions_rows_iata = []
                transactions_rows_gsa = []
                total_due_to_airlinepros = []
                total_due_to_airlinepros_negative = []
                acms_rows = []
                acms_rows_total = []

                total_ap_acm = 0.00
                weekly_sales_total = 0.00
                total_cc = 0.00
                total_ca = 0.00

                net_sales_weekly_total = 0.00
                net_sales_weekly_total_iata = 0.00
                net_sales_weekly_total_gsa = 0.00
                pending_deduction_total = 0.00
                weekly_total_cash_disbursement_total = 0.00
                arc_deductions_total = 0.00
                total_commission_var_p = 0
                total_cp_total = 0
                total_fare_total = 0

                cash_t = []
                amex_t = []
                visa_t = []
                mstr_t = []
                other_t = []

                for t in transactions:
                    
                    transactions_headers.append(t.get('report__report_period__ped'))

                    total_tax_yq_yr = Charges.objects.select_related('transaction').filter(type__in=['YQ', 'YR'],
                                                                                           transaction__report__report_period__ped=t.get(
                                                                                               'report__report_period__ped'),
                                                                                           transaction__report__airline=airline).aggregate(
                        total_tax_yq_yr=Coalesce(Sum('amount'), V(0))).get('total_tax_yq_yr')

                    total_tax_cp_mf = Charges.objects.select_related('transaction').filter(type__in=['CP', 'MF'],
                                                                                           transaction__report__report_period__ped=t.get(
                                                                                               'report__report_period__ped'),
                                                                                           transaction__report__airline=airline).aggregate(
                        total_tax_cp_mf=Coalesce(Sum('amount'), V(0))).get('total_tax_cp_mf')



                    penalty_charge = Charges.objects.select_related('transaction').filter(type__in=['CP'],
                                                                                           transaction__report__report_period__ped=t.get(
                                                                                               'report__report_period__ped'),
                                                                                           transaction__report__airline=airline).aggregate(
                        total_tax_cp_mf=Coalesce(Sum('amount'), V(0))).get('total_tax_cp_mf')



                    if t.get('total_fare_amount'):

                        print(t.get('total_fare_amount'),"//////////")
                        total_fare_amount_var = t.get('total_fare_amount')
                        total_fare_total += total_fare_amount_var
                    else:
                        total_fare_amount_var = 0.00

                    if t.get('cp'):
                        cp_var = t.get('cp')
                    else:
                        cp_var = 0.00
                    total_cp_total += cp_var
                    total_fare_amount_var = total_fare_amount_var + penalty_charge

                    total_ap_acm_var = 0.00
                    if t.get('total_ap_acm'):
                        total_ap_acm_var = t.get('total_ap_acm')


                    total_fare_amount_var = total_fare_amount_var - total_ap_acm_var
                    net_sales_weekly = total_fare_amount_var

                    total_tax_var = Taxes.objects.select_related('transaction').filter(
                        transaction__report__report_period__ped=t.get('report__report_period__ped'),
                        transaction__report__airline=airline).aggregate(
                        total_tax=Coalesce(Sum('amount'), V(0))).get('total_tax')



                    if t.get('total_commission'):
                        total_commission_var = t.get('total_commission') * -1
                        
                    else:
                        total_commission_var = 0.00
                    total_commission_var_p += total_commission_var



                    weekly_sales = total_fare_amount_var + total_tax_var + total_commission_var + total_tax_yq_yr + value_check(
                        t.get('total_ap_acm'))

                    remittance = t.get('report__report_period__remittance_date')

                    try:
                        disb = Disbursement.objects.get(report_period__ped=t.get('report__report_period__ped'),
                                                        airline=airline)
                        cash_disbursement = disb.bank7
                        arc_deductions = disb.arc_deduction + disb.arc_fees + disb.arc_reversal + disb.arc_tot
                        weekly_total_cash_disbursement = disb.arc_net_disb
                    except Exception as e:
                        cash_disbursement = 0.00
                        arc_deductions = 0.00
                        weekly_total_cash_disbursement = 0.00

                    try:
                        pending_deduction = Deduction.objects.filter(
                            report__report_period__ped=t.get('report__report_period__ped'), report__airline=airline,
                            pending=True).aggregate(pending_deduction=Coalesce(Sum('amount'), V(0))).get(
                            'pending_deduction')
                    except Exception as e:
                        pending_deduction = 0.00

                    arc_deductions_total = arc_deductions_total + arc_deductions

                    weekly_total_cash_disbursement_total = weekly_total_cash_disbursement_total + weekly_total_cash_disbursement
                    card_disbursement = weekly_sales - cash_disbursement


                    pending_deduction_total = pending_deduction_total + pending_deduction


                    ped_total = total_fare_amount_var
                    pen_val = 0
                    if t.get("pen_value"):
                        pen_val = float(t.get("pen_value"))
                    if not total_tax_cp_mf:
                        total_tax_cp_mf = 0
                    calculated_fare_amount = "{:.2f}".format(
                        total_fare_amount_var + total_tax_cp_mf - total_commission_var)
                    transactions_rows.append(
                        {"fare": total_fare_amount_var, "tax": total_tax_var, "comm": total_commission_var,
                         "calculated_fare_amount": calculated_fare_amount,
                         "tax_yq_yr": total_tax_yq_yr, "weekly_sales_total": weekly_sales,
                         "total_ca": cash_disbursement, "total_cc": card_disbursement,
                         "total_ap_acm": t.get("total_ap_acm"), "remittance": remittance,
                         "week": t.get('report__report_period__week'), "pending_deduction": pending_deduction,
                         'net_sales_weekly': net_sales_weekly, 'ped_total': ped_total, 'arc_deductions': arc_deductions,
                         'weekly_total_cash_disbursement': weekly_total_cash_disbursement})

                    net_sales_weekly_total = net_sales_weekly_total + net_sales_weekly

                    if t.get('total_ap_acm'):
                        total_ap_acm = total_ap_acm + t.get('total_ap_acm')

                    weekly_sales_total = weekly_sales_total + weekly_sales

                    if card_disbursement:
                        total_cc = total_cc + card_disbursement
                    if cash_disbursement:
                        total_ca = total_ca + cash_disbursement


                    transactions_rows_iata.append((net_sales_weekly * iata_coordination_fee) / 100)
                    transactions_rows_gsa.append((net_sales_weekly * gsa_commision) / 100)

                    net_sales_weekly_total_iata = net_sales_weekly_total_iata + (
                        (net_sales_weekly * iata_coordination_fee) / 100)
                    net_sales_weekly_total_gsa = net_sales_weekly_total_gsa + (
                        (net_sales_weekly * gsa_commision) / 100)

                    total_due_to_airlinepros.append(
                        ((net_sales_weekly * iata_coordination_fee) / 100) + ((net_sales_weekly * gsa_commision) / 100))
                    total_due_to_airlinepros_negative.append(((((net_sales_weekly * iata_coordination_fee) / 100) + (
                        (net_sales_weekly * gsa_commision) / 100))) * -1)
                

            else:
                
                month_date = datetime.datetime.strptime(month_year, '%B %Y')

                commission_history = CommissionHistory.objects.filter(airline=airline, type='I',
                                                                      from_date__lte=month_date)
                if commission_history.exists():
                    iata_coordination_fee = commission_history.order_by('-from_date').first().rate

                commission_history = CommissionHistory.objects.filter(airline=airline, type='G',
                                                                      from_date__lte=month_date)
                if commission_history.exists():
                    gsa_commision = commission_history.order_by('-from_date').first().rate

                transactions = Transaction.objects.exclude(transaction_type__startswith='SP',
                                                           agency__agency_no='6999001').select_related('report',
                                                                                                   'report__airline',
                                                                                                       'report__report_period').filter(
                    report__airline=airline,
                    report__report_period__month=month,
                    report__report_period__year=year).values(
                    'report__report_period__ped', 'report__report_period__week',
                    'report__report_period__remittance_date').annotate(
                    dcount=Count('report__report_period__ped'),
                    total_transaction_amount=Sum('transaction_amount'),
                    total_fare_amount=Sum('fare_amount'),
                    total_commission=Sum('std_comm_amount'),
                    tax_on_commission=Sum('tax_on_comm'),
                    report__cc=Sum('cc'),
                    report__ca=Sum('ca'),
                    ep=Sum('ep'),
                    Balance_Payable=Sum('balance'),
                    pen_value=Sum('pen'),
                    cp=Sum('pen', filter=Q(report__airline=airline, pen_type='CP')),
                    total_tax=Subquery(sub_tax, output_field=FloatField()),
                    total_adm=Subquery(total_adm, output_field=FloatField()),
                    total_tax_yq_yr=Subquery(sub_tax_yq_yr,
                                             output_field=FloatField()),
                    total_tax_cp_mf=Subquery(sub_tax_cp_mf, output_field=FloatField())
                ).order_by(
                    'report__report_period__ped')
                

                transactions_headers = []
                transactions_rows = []
                transactions_rows_iata = []
                transactions_rows_gsa = []
                total_due_to_airlinepros = []
                total_due_to_airlinepros_negative = []
                acms_rows = []
                acms_rows_total = []

                total_ap_acm = 0.00
                weekly_sales_total = 0.00
                total_cc = 0.00
                total_ca = 0.00

                net_sales_weekly_total = 0.00
                net_sales_weekly_total_iata = 0.00
                net_sales_weekly_total_gsa = 0.00
                pending_deduction_total = 0.00
                weekly_total_cash_disbursement_total = 0.00
                arc_deductions_total = 0.00
                total_ap_acm_var = 0.00
                if not transactions:
                    
                    for c_t in range(0, 4):
                        t = {}
                        ap_acm = {}
                        month_day = self.get_all_sundays(month, year)[c_t]
                        transactions_headers.append(month_day)

                        total_tax_yq_yr = 0.00
                        total_tax_cp_mf = 0.00
                        total_fare_amount_var = 0.00
                        total_commission_var = 0.00
                        total_tax_var = 0.00
                        ap_acm_comm = 0.00
                        total_commission_var = 0.00
                        cp_var = 0.00
                        ap_acm_cp = 0.00
                        cp_var = 0.00


                        remittance = month_day

                        cash_disbursement = 0.00

                        pending_deduction_total = 0.00

                        net_sales_weekly = 0.00

                        weekly_sales = 0.00
                        pen_val = 0
                        if t.get("pen_value"):
                            pen_val = float(t.get("pen_value"))
                        calculated_fare_amount = "{:.2f}".format(total_fare_amount_var + pen_val - total_commission_var)
                        transactions_rows.append(
                            {"fare": total_fare_amount_var, "tax": t.get('total_tax'),
                             "calculated_fare_amount": calculated_fare_amount,
                             "comm": (abs(total_commission_var) * -1),
                             "tax_yq_yr": total_tax_yq_yr, "tax_cp_mf": total_tax_cp_mf,
                             "weekly_sales_total": weekly_sales,
                             "total_ca": cash_disbursement, "total_cc": t.get("report__cc"),
                             "total_ap_acm": total_ap_acm_var, "remittance": remittance,
                             "week": c_t + 1, "pending_deduction": 0.00,
                             'net_sales_weekly': net_sales_weekly, 'ped_total': total_fare_amount_var})

                        net_sales_weekly_total = net_sales_weekly_total + net_sales_weekly

                        total_ap_acm = 0.00

                        weekly_sales_total = 0.00

                        total_cc = 0.00
                        total_ca = 0.00

                        transactions_rows_iata.append((net_sales_weekly * iata_coordination_fee) / 100)
                        transactions_rows_gsa.append((net_sales_weekly * gsa_commision) / 100)

                        net_sales_weekly_total_iata = net_sales_weekly_total_iata + (
                            (net_sales_weekly * iata_coordination_fee) / 100)
                        net_sales_weekly_total_gsa = net_sales_weekly_total_gsa + (
                            (net_sales_weekly * gsa_commision) / 100)

                        total_due_to_airlinepros.append(
                            ((net_sales_weekly * iata_coordination_fee) / 100) + (
                                (net_sales_weekly * gsa_commision) / 100))
                        total_due_to_airlinepros_negative.append(
                            ((((net_sales_weekly * iata_coordination_fee) / 100) + (
                                (net_sales_weekly * gsa_commision) / 100))) * -1)
                
                Total_transaction_amount = 0
                Total_fare = 0
                Total_tax = 0
                Total_tax_yq_yr = 0
                Total_Penalties = 0
                Total_Easy_pay = 0
                Total_comm = 0
                Total_tax_on_commission = 0
                Total_Balance_Payable = 0
                Total_cc = 0
                Total_ca = 0
                Total_amd = 0
                ACSA_Fee = 0
                
               
                
                
                
                for t in transactions:
                    
                    transactions_headers.append(t.get('report__report_period__ped'))

                    if t.get("total_tax_yq_yr"):
                        total_tax_yq_yr = t.get("total_tax_yq_yr")
                    else:
                        total_tax_yq_yr = 0.00

                    if t.get("total_tax_cp_mf"):
                        total_tax_cp_mf = t.get("total_tax_cp_mf")
                    else:
                        total_tax_cp_mf = 0.00

                    ap_acm = Transaction.objects.filter(report__report_period__ped=t.get('report__report_period__ped'),
                                                        report__airline=airline, transaction_type__startswith='ACM',
                                                        agency__agency_no='6998051').aggregate(
                        total=Sum('fare_amount', output_field=FloatField()),
                        total_comm=Sum('std_comm_amount', output_field=FloatField()),
                        total_cp=Sum('pen', output_field=FloatField()))

                    total_cp = 0
                    for ac in Transaction.objects.filter(report__report_period__ped=t.get('report__report_period__ped'),
                                                         report__airline=airline, transaction_type__startswith='ACM',
                                                         agency__agency_no='6998051'):
                        for chrge in ac.charges_set.filter(type__in=['CP', 'MF']):
                            total_cp += int(chrge.amount)
                    
                    if t.get('total_fare_amount'):
                        total_ap_acm_var = 0.00
                        if ap_acm.get('total'):
                            total_ap_acm_var = ap_acm.get('total')
                        total_fare_amount_var = t.get('total_fare_amount') + total_cp
                    else:
                        total_fare_amount_var = 0.00
                    if t.get('total_tax'):
                        total_tax_var = t.get('total_tax')
                    else:
                        total_tax_var = 0.00
                        
                    if t.get('total_transaction_amount'):
                        total_transaction_amount_var = t.get('total_transaction_amount')
                    else:
                        total_transaction_amount_var = 0.00    
                    
                    if t.get('Balance_Payable'):
                        balance_Payable_var = t.get('Balance_Payable')
                    else:
                        balance_Payable_var = 0.00

                    if t.get('tax_on_commission'):
                        tax_on_commission_var = t.get('tax_on_commission')
                    else:
                        tax_on_commission_var = 0.00

                    if t.get('ep'):
                        Easy_pay_var = t.get('ep')
                    else:
                        Easy_pay_var = 0.00

                    if t.get('total_adm'):
                        total_adm = t.get('total_adm')
                    else:
                        total_adm = 0.00

                    total_commission_var = 0.00
                    if t.get('total_commission'):
                        total_commission_var = t.get('total_commission')

                    ap_acm_comm = 0.00
                    if ap_acm.get('total_comm'):
                        ap_acm_comm = ap_acm.get('total_comm')

                    total_commission_var = total_commission_var - ap_acm_comm

                    cp_var = 0.00
                    if t.get('cp'):
                        cp_var = t.get('cp')

                    ap_acm_cp = 0.00
                    if ap_acm.get('total_cp'):
                        ap_acm_cp = ap_acm.get('total_cp')

                    cp_var = cp_var - ap_acm_cp


                    remittance = t.get('report__report_period__remittance_date')



                    cash_disbursement = t.get("report__ca") - t.get('total_commission')

                    pending_deduction = value_check(t.get("report__ca")) + value_check(t.get('total_adm')) - (
                        value_check(t.get('total_commission'))) - value_check(t.get('total_ap_acm'))

                    pending_deduction_total = pending_deduction_total + pending_deduction

                    try:
                        country_obj = Country.objects.get(id=country)
                        if country_obj.name.lower() in ['canada', 'united states']:
                            net_sales_weekly = total_fare_amount_var - total_commission_var + total_tax_cp_mf
                        else:
                            net_sales_weekly = t.get("total_fare_amount")
                    except:
                        net_sales_weekly = total_fare_amount_var - total_commission_var + total_tax_cp_mf


                    weekly_sales = net_sales_weekly + total_tax_var + total_tax_yq_yr
                    pen_val = 0
                    if t.get("pen_value"):
                        pen_val = float(t.get("pen_value"))
                    calculated_fare_amount = "{:.2f}".format(total_fare_amount_var + pen_val - total_commission_var)
                    
                    Total_transaction_amount = Total_transaction_amount + total_transaction_amount_var
                    Total_Balance_Payable = Total_Balance_Payable + balance_Payable_var
                    Total_tax_on_commission = Total_tax_on_commission + tax_on_commission_var
                    Total_Easy_pay = Total_Easy_pay + Easy_pay_var
                    Total_comm = Total_comm + (abs(total_commission_var))
                    Total_Penalties = Total_Penalties + total_tax_cp_mf
                    Total_tax_yq_yr = Total_tax_yq_yr + total_tax_yq_yr
                    if not t.get('total_tax'):
                        Total_tax = Total_tax + 0
                    else:
                        Total_tax = Total_tax + t.get('total_tax')   
                    Total_fare = Total_fare + total_fare_amount_var
                    Total_amd = Total_amd + total_adm
                    Total_ca = Total_ca + (balance_Payable_var+(abs(total_commission_var)))
                    Total_cc = Total_cc + t.get("report__cc")
                    
                    transactions_rows.append(
                        {"fare": total_fare_amount_var, "Transaction_amount":total_transaction_amount_var, 
                         "Balance_Payable": balance_Payable_var, "tax_on_commission": tax_on_commission_var , 
                         "tax": t.get('total_tax'), "Easy_pay": Easy_pay_var, 
                         "calculated_fare_amount": calculated_fare_amount, "tax_yq_yr": total_tax_yq_yr,
                         "comm": (abs(total_commission_var)), "Total_Penalties" :total_tax_cp_mf,
                         "tax_yq_yr": total_tax_yq_yr, "tax_cp_mf": total_tax_cp_mf, "weekly_sales_total": weekly_sales,
                         "total_ca": (balance_Payable_var+(abs(total_commission_var))), "total_cc": t.get("report__cc"),
                         "total_ap_acm": total_ap_acm_var, "remittance": remittance,
                         "week": t.get('report__report_period__week'), "pending_deduction": pending_deduction,
                         'net_sales_weekly': net_sales_weekly, 'ped_total': total_fare_amount_var, 
                         "Total_transaction_amount": Total_transaction_amount, "Total_Balance_Payable" :Total_Balance_Payable,
                         "Total_tax_on_commission": Total_tax_on_commission, "Total_Easy_pay": Total_Easy_pay,
                         "Total_comm": Total_comm, "Total_Penalties": Total_Penalties, "Total_tax_yq_yr": Total_tax_yq_yr,
                         "Total_tax":Total_tax, "Total_fare": Total_fare, "Total_ca": Total_ca, "Total_cc": Total_cc ,"Total_amd":Total_amd})

                    net_sales_weekly_total = net_sales_weekly_total + net_sales_weekly
                    if total_ap_acm_var:
                        total_ap_acm = total_ap_acm + total_ap_acm_var

                    weekly_sales_total = weekly_sales_total + weekly_sales

                    if t.get('report__cc'):
                        total_cc = total_cc + t.get('report__cc')
                    if cash_disbursement:
                        total_ca = total_ca + cash_disbursement



                    transactions_rows_iata.append((net_sales_weekly * iata_coordination_fee) / 100)
                    transactions_rows_gsa.append((net_sales_weekly * gsa_commision) / 100)

                    net_sales_weekly_total_iata = net_sales_weekly_total_iata + (
                        (net_sales_weekly * iata_coordination_fee) / 100)
                    net_sales_weekly_total_gsa = net_sales_weekly_total_gsa + (
                        (net_sales_weekly * gsa_commision) / 100)

                    total_due_to_airlinepros.append(
                        ((net_sales_weekly * iata_coordination_fee) / 100) + ((net_sales_weekly * gsa_commision) / 100))
                    total_due_to_airlinepros_negative.append(((((net_sales_weekly * iata_coordination_fee) / 100) + (
                        (net_sales_weekly * gsa_commision) / 100))) * -1)
                   
            TTL_BSP_SALES_Ex_Tax = Total_cc + Total_ca + Total_Easy_pay - Total_tax
            if country == '12' or  country == '13':    
                if airline == '126' or airline == '132':
                    ACSA_Fee = TTL_BSP_SALES_Ex_Tax * (3/100)
                else:
                    ACSA_Fee = TTL_BSP_SALES_Ex_Tax * (iata_coordination_fee/100)
            else:
                ACSA_Fee = TTL_BSP_SALES_Ex_Tax * (iata_coordination_fee/100)
            month_year = self.request.GET.get('month_year', '')
            airline = self.request.GET.get('airline', '')
            country = self.request.session.get('country')
            self.is_arc_var = is_arc(self.request.session.get('country'))

            # adms = []
            
            # if airline or month_year:
            #     if not self.is_arc_var:
            #         trans_type = 'TKTT'
            #     else:
            #         trans_type = 'TKT'

            #     qs = Transaction.objects.filter(transaction_type=trans_type)
            #     adm_objs = AgencyDebitMemo.objects.select_related('transaction').filter(
            #         transaction__transaction_type=trans_type)
                
            #     if airline:
            #         qs = qs.filter(report__airline=airline)
            #         #

            #         adm_objs = adm_objs.filter(transaction__report__airline=airline)
            #     if month_year:
            #         month = datetime.datetime.strptime(month_year, '%B %Y').month or ''
            #         year = datetime.datetime.strptime(month_year, '%B %Y').year or ''
            #         month_date = datetime.datetime.strptime(month_year, '%B %Y')
            #         if month and year:
            #             qs = qs .filter(report__report_period__month=month,
            #                         report__report_period__year=year)
            #             #

            #             adm_objs = adm_objs.filter(transaction__report__report_period__month=month,
            #                                 transaction__report__report_period__year=year)
            #             #

            #             if qs.count() + adm_objs.count() > 10000:
            #                 # 10000:
            #                 base_url = self.request.scheme + '://' + self.request.get_host()
            #                 excel_adm_report(country, month_year, airline, self.request.user.email, base_url)
            #                 messages.add_message(self.request, messages.WARNING,
            #                                     'Excel file generation is taking more time than expected. You will receive an email with the link to download the file once it is done.')
            #                 return HttpResponseRedirect('/reports/adm/?' + self.request.META.get('QUERY_STRING'))

            #             adms_list = adm_objs.values_list('transaction__ticket_no', flat=True)
            #             # ---------------------------------------------------------------------------


            #             allowed_commission_rate = 0.00
            #             Total_amd = 0
            #             for obj in qs:
            #                 if obj.ticket_no not in adms_list  :
            #                     allowed_commission_rate = 0.00

            #                     commission_history = CommissionHistory.objects.filter(airline=airline, type='M',
            #                                                                         from_date__lte=obj.report.report_period.ped)
            #                     if commission_history.exists():
            #                         allowed_commission_rate = commission_history.order_by('-from_date').first().rate

            #                     try:
            #                         taken_commission_rate = obj.std_comm_rate
            #                         commission_rate_diff = allowed_commission_rate - float(taken_commission_rate)
            #                         cobl_amount = obj.fare_amount
            #                         if cobl_amount:
            #                             if commission_rate_diff < 0:
            #                                 # ADM
            #                                 adm_amount = (abs(commission_rate_diff) * cobl_amount) / 100
            #                                 Total_amd = Total_amd + adm_amount
                                            
            #                                 comment = "Commission deducted " + str(
            #                                     taken_commission_rate) + "%. Carrier authorized " + str(
            #                                     allowed_commission_rate) + "%"
            #                                 allowed_commission_amount = (cobl_amount * allowed_commission_rate) / 100
            #                                 adms.append({
            #                                     'agency_no': obj.agency.agency_no,
            #                                     'trade_name': obj.agency.trade_name,
            #                                     'ticket_no': obj.ticket_no,
            #                                     'issue_date': obj.issue_date,
            #                                     'fare_amount': obj.fare_amount,
            #                                     'std_comm_amount': obj.std_comm_amount,
            #                                     'std_comm_rate': obj.std_comm_rate,
            #                                     'allowed_commission_amount': allowed_commission_amount,
            #                                     'amount': adm_amount,
            #                                     'comment': comment

            #                                 })
            #                     except Exception as e:
            #                         pass
                            
                        
            Total_ACM = sum(total_due_to_airlinepros_negative)
            
            total_Transaction = Transaction.objects.filter(report__airline = airline_data,report__report_period__month = month, report__country = country ).exclude(transaction_type="+RTDN")

            li = []
            for i in total_Transaction:
                li.append(i.ticket_no)
            Total_Transaction = len(set(li))
            
            
            
            # sheet-start
            wb = xlwt.Workbook(style_compression=2)
            ws = wb.add_sheet('BSP-SalesSummary',cell_overwrite_ok=True)


            ws.col(0).width = (256 * 12)
            row_num = 0
            ssd_new=xlwt.easyxf("align: wrap yes, vert centre, horiz right;font: name Arial, height 180, bold True;border: left medium,right medium,top medium,bottom medium;pattern: pattern solid, fore_colour red")

            left_normal_bold = xlwt.easyxf(
                "align: wrap yes, vert centre, horiz left;font: name Arial, bold True, height 180;border: left thin,right thin,top thin,bottom thin")

            bold_center = xlwt.easyxf("font: name Arial, bold True, height 280; align: horiz center")
            wrap_data = xlwt.easyxf("align: wrap yes, vert centre, horiz center;font: name Arial,height 180")
            center_normal = xlwt.easyxf(
                "align: wrap yes, vert centre, horiz center;font: name Arial, height 180;border: left thin,right thin,top thin,bottom thin")
            
            normal_left = xlwt.easyxf(
                "align: wrap yes, vert centre, horiz center;font: name Arial, height 180;border: left thin")
            normal_right = xlwt.easyxf(
                "align: wrap yes, vert centre, horiz center;font: name Arial, height 180;border: right thin")
            
            center_normal_right = xlwt.easyxf(
                "align: wrap yes, vert centre, horiz right;font: name Arial, height 180;border: left thin,right thin,top thin,bottom thin")
            center_normal_border = xlwt.easyxf(
                "align: wrap yes, vert centre, horiz center;font: name Arial, height 180;border: left medium,right medium,top medium,bottom medium")
            center_normal_bold = xlwt.easyxf(
                "align: wrap yes, vert centre, horiz center;font: name Arial, bold True, height 180;border: left medium,right medium,top medium,bottom medium")
            
            center_normal_border_right = xlwt.easyxf(
                "align: wrap yes, vert centre, horiz right;font: name Arial, height 180, bold True;border: left medium,right medium,top medium,bottom medium")

            center_normal_border_yellow_center = xlwt.easyxf(
                "align: wrap yes, vert centre, horiz centre;font: name Arial, height 180, bold True;border: left medium,right medium,top medium,bottom medium;pattern: pattern solid, fore_colour yellow")

            left_normal_border_yellow_center = xlwt.easyxf(
                "align: wrap yes, vert centre, horiz left;font: name Arial, height 180, bold True;border: left medium,right medium,top medium,bottom medium;pattern: pattern solid, fore_colour yellow")


            center_normal_border_yellow_right = xlwt.easyxf(
                "align: wrap yes, vert centre, horiz right;font: name Arial, height 180, bold True;border: left medium,right medium,top medium,bottom medium;pattern: pattern solid, fore_colour yellow")
            center_normal_border_date = xlwt.easyxf(
                "align: wrap yes, vert centre, horiz center;font: name Arial, height 180, bold True;border: left medium,right medium,top medium,bottom medium")
            center_normal_border_color = xlwt.easyxf(
                "align: wrap yes, vert centre, horiz center;font: name Arial, height 180;border: left medium,right medium,top medium,bottom medium;pattern: pattern solid, fore_colour yellow")

            center_normal_border_color_right = xlwt.easyxf(
                "align: wrap yes, vert centre, horiz right;font: name Arial, height 180, bold True;border: left medium,right medium,top medium,bottom medium;pattern: pattern solid, fore_colour yellow")
            center_normal_color = xlwt.easyxf(
                "align: wrap yes, vert centre, horiz center;font: name Arial, height 180;pattern: pattern solid, fore_colour yellow;border: left thin,right thin,top thin,bottom thin")
            center_normal_color_right = xlwt.easyxf(
                "align: wrap yes, vert centre, horiz right;font: name Arial, height 180;pattern: pattern solid, fore_colour yellow;border: left thin,right thin,top thin,bottom thin")
            center_normal_red_right = xlwt.easyxf(
                "align: wrap yes, vert centre, horiz right;font: name Arial, height 180, colour red;pattern: pattern solid, fore_colour yellow;border: left thin,right thin,top thin,bottom thin")
            center_normal_color_date = xlwt.easyxf(
                "align: wrap yes, vert centre, horiz right;font: name Arial, height 180;pattern: pattern solid, fore_colour yellow;border: left thin,right thin,top thin,bottom thin")
            ws.row(0).height_mismatch = True
            ws.row(row_num).height = 19 * 21
            ws.write_merge(row_num, 0, 0, 11, airline_name.upper(), bold_center)
            row_num = row_num + 1
            ws.row(row_num).height = 19 * 21
            ws.write_merge(row_num, 1, 0, 11, "BSP-Sales Summary Reports".upper(), bold_center)
            row_num = row_num + 1
            ws.row(row_num).height = 19 * 21
            date_month_row = month_year.split(' ')[0][:3] + '-' + month_year.split(' ')[1]
            ws.write_merge(row_num, 2, 0, 11, date_month_row, bold_center)
            row_num = 3
            col_num_1 = 1
            col_num_2 = 4 
            col_num_3 = 9
            col_num_4 = 10
            col_num_5 = 5
            ws.row(row_num).height = (20*22)
            ws.col(col_num_1).width = (200*20)
            ws.col(col_num_2).width = (200*20)
            ws.col(col_num_3).width = (200*20)
            ws.col(col_num_4).width = (200*20)
            ws.col(col_num_5).width = (200*20)
            ws.write_merge(row_num, 4, 0, 0, "BILLING PERIOD", center_normal_bold)
            ws.write_merge(row_num, 4, col_num_1, 1, "TRANSACTION AMOUNT", center_normal_bold)
            ws.write_merge(3, 4, 2, 2, "FARE", center_normal_bold)
            ws.write_merge(3, 4, 3, 3, "TAX", center_normal_bold)
            ws.write_merge(row_num, 4, col_num_2, 4, "TAXES,FEES AND CHARGES", center_normal_bold)
            ws.write_merge(row_num, 4, col_num_5, 5, "PENALTIES", center_normal_bold)
            ws.write_merge(3, 4, 6, 6, "CASH", center_normal_bold)
            ws.write_merge(row_num, 4, 7, 7, "CREDIT CARD", center_normal_bold)
            ws.write_merge(row_num, 4, 8, 8, "EASY PAY", center_normal_bold)
            ws.write_merge(row_num, 4, col_num_3, 9, "COMMISSION", center_normal_bold)
            ws.write_merge(row_num, 4, col_num_4, 10, "TAX ON COMMISSION", center_normal_bold)
            ws.write_merge(row_num, 4, 11, 11, "BALANCE PAYABLE", center_normal_bold)
            row_num_right = [5,6,7,8,9]
            for row in row_num_right:
                ws.write(row, 11, " ", normal_right)
            row_num_left = [5,6,7,8,9]
            for row in row_num_left:
                ws.write(row, 0, " ", normal_left)

            row = 5
            
            for ele in transactions_rows:
                ws.write(row, 0, "PERIOD"+" "+str(ele.get("week")),center_normal)
                if ele.get("Transaction_amount") is None:
                    ws.write(row, 1, (ele.get("Transaction_amount")),center_normal)
                else:
                    ws.write(row, 1, round(ele.get("Transaction_amount"),2),center_normal)
                
                if ele.get("fare") is None:
                    ws.write(row, 2, (ele.get("fare")),center_normal)
                else:
                    ws.write(row, 2, round(ele.get("fare"),2),center_normal)
                
                if ele.get("tax") is None:
                    ws.write(row, 3, (ele.get("tax")),center_normal)
                else:
                    ws.write(row, 3, round(ele.get("tax"),2),center_normal)
                    
                if ele.get("tax_yq_yr") is None:
                    ws.write(row, 4, (ele.get("tax_yq_yr")),center_normal)
                else:    
                   ws.write(row, 4, round(ele.get("tax_yq_yr"),2),center_normal)
                
                if ele.get("tax_cp_mf") is None:
                    ws.write(row, 5, (ele.get("tax_cp_mf")),center_normal)
                else:    
                  ws.write(row, 5, round(ele.get("tax_cp_mf"),2),center_normal)
                
                if ele.get("total_ca") is None:
                    ws.write(row, 6, (ele.get("total_ca")),center_normal)
                else:    
                  ws.write(row, 6, round(ele.get("total_ca"),2),center_normal)
                  
                if ele.get("total_cc") is None:
                    ws.write(row, 7, (ele.get("total_cc")),center_normal)
                else:  
                    ws.write(row, 7, round(ele.get("total_cc"),2),center_normal)
                    
                if ele.get("Easy_pay") is None:
                    ws.write(row, 8, (ele.get("Easy_pay")),center_normal)
                else:
                    ws.write(row, 8, round(ele.get("Easy_pay"),2),center_normal)
                    
                if ele.get("comm") is None:
                    ws.write(row, 9, (ele.get("comm")),center_normal)
                else:
                    ws.write(row, 9, round(ele.get("comm"),2),center_normal)
                    
                if ele.get("tax_on_commission") is None:
                    ws.write(row, 10, (ele.get("tax_on_commission")),center_normal)
                else:
                    ws.write(row, 10,round(ele.get("tax_on_commission"),2),center_normal)
                    
                if ele.get("Balance_Payable") is None:
                    ws.write(row, 11, (ele.get("Balance_Payable")),center_normal)
                else:
                    ws.write(row, 11,round(ele.get("Balance_Payable"),2),center_normal)
                
                
                    
                row_num = 10
                
                ws.row(row_num).height = (20*22)
                ws.write(row_num, 0, ("GRAND TOTAL"), center_normal_border_yellow_center)
                
                if ele.get("Total_transaction_amount") is None:
                    ws.write(row_num, 1, (ele.get("Total_transaction_amount")), center_normal_border_yellow_center)
                else:
                    ws.write(row_num, 1, round(ele.get("Total_transaction_amount"),2), center_normal_border_yellow_center)
                
                if ele.get("Total_fare") is None:
                    ws.write(row_num, 2, (ele.get("Total_fare")), center_normal_border_yellow_center)
                else:
                    ws.write(row_num, 2, round(ele.get("Total_fare"),2), center_normal_border_yellow_center)
                
                if ele.get("Total_tax") is None:
                    ws.write(row_num, 3, (ele.get("Total_tax")), center_normal_border_yellow_center)
                else:
                    ws.write(row_num, 3, round(ele.get("Total_tax"),2), center_normal_border_yellow_center)
                
                if ele.get("Total_tax_yq_yr") is None:
                    ws.write(row_num, 4, (ele.get("Total_tax_yq_yr")), center_normal_border_yellow_center)
                else:
                    ws.write(row_num, 4, round(ele.get("Total_tax_yq_yr"),2), center_normal_border_yellow_center)
                
                if ele.get("Total_Penalties") is None:
                    ws.write(row_num, 5, (ele.get("Total_Penalties")), center_normal_border_yellow_center)
                else:
                    ws.write(row_num, 5, round(ele.get("Total_Penalties"),2), center_normal_border_yellow_center)
                
                if ele.get("Total_ca") is None:
                    ws.write(row_num, 6, (ele.get("Total_ca")), center_normal_border_yellow_center)
                else:
                    ws.write(row_num, 6, round(ele.get("Total_ca"),2), center_normal_border_yellow_center)
                    
                if ele.get("Total_cc") is None:
                    ws.write(row_num, 7, (ele.get("Total_cc")), center_normal_border_yellow_center)
                else:
                    ws.write(row_num, 7, round(ele.get("Total_cc"),2), center_normal_border_yellow_center)
                    
                if ele.get("Total_Easy_pay") is None:
                    ws.write(row_num, 8, (ele.get("Total_Easy_pay")), center_normal_border_yellow_center)
                else:
                    ws.write(row_num, 8, round(ele.get("Total_Easy_pay"),2), center_normal_border_yellow_center)
                
                if ele.get("Total_comm") is None:
                    ws.write(row_num, 9, (ele.get("Total_comm")), center_normal_border_yellow_center)
                else:
                    ws.write(row_num, 9, round(ele.get("Total_comm"),2), center_normal_border_yellow_center)
                    
                if ele.get("Total_tax_on_commission") is None:
                    ws.write(row_num, 10, (ele.get("Total_tax_on_commission")), center_normal_border_yellow_center)
                else:
                    ws.write(row_num, 10, round(ele.get("Total_tax_on_commission"),2), center_normal_border_yellow_center)
                
                if ele.get("Total_Balance_Payable") is None:
                    ws.write(row_num, 11, (ele.get("Total_Balance_Payable")), center_normal_border_yellow_center)
                else:
                    ws.write(row_num, 11, round(ele.get("Total_Balance_Payable"),2), center_normal_border_yellow_center)
                    
                
                    
                    
                row += 1
                
                
            
            
            
            ws.write_merge(14, 15, 0, 3, "NUMBER OF TRANSACTIONS", left_normal_bold)
            ws.write_merge(14, 15, 4, 5, "",center_normal)
            if Total_Transaction is None:
               ws.write_merge(14, 15, 6, 6, (Total_Transaction), center_normal) 
            else:
                ws.write_merge(14, 15, 6, 6, round(Total_Transaction,2), center_normal)
            
            ws.write_merge(16, 17, 0, 3, "TOTAL BSP SALES CASH/CHECK", left_normal_bold)
            ws.write_merge(16, 17, 4, 5, (currency_code),center_normal)
            if Total_ca is None:
               ws.write_merge(16, 17, 6, 6, (Total_ca), center_normal) 
            else:
                ws.write_merge(16, 17, 6, 6, round(Total_ca,2), center_normal)
            
            ws.write_merge(18, 19, 0, 3, "TOTAL BSP SALES CREDIT CARD", left_normal_bold)
            ws.write_merge(18, 19, 4, 5, (currency_code), center_normal)
            if Total_cc is None:
               ws.write_merge(18, 19, 6, 6, (Total_cc), center_normal) 
            else:
                ws.write_merge(18, 19, 6, 6, round(Total_cc,2), center_normal)
            
            ws.write_merge(20, 21, 0, 3, "TOTAL BSP SALES EASY PAY",left_normal_bold)
            ws.write_merge(20, 21, 4, 5, (currency_code), center_normal)
            if Total_Easy_pay is None:
               ws.write_merge(20, 21, 6, 6, (Total_Easy_pay), center_normal) 
            else:
                ws.write_merge(20, 21, 6, 6, round(Total_Easy_pay,2), center_normal)
            
            ws.write_merge(22, 23, 0, 3, "TOTAL YQ/YR",left_normal_bold)
            ws.write_merge(22, 23, 4, 5, (currency_code), center_normal)
            if Total_tax_yq_yr is None:
               ws.write_merge(22, 23, 6, 6, (Total_tax_yq_yr), center_normal) 
            else:
                ws.write_merge(22, 23, 6, 6, round(Total_tax_yq_yr,2), center_normal)
            
            ws.write_merge(24, 25, 0, 3, "TOTAL PENALTY",left_normal_bold)
            ws.write_merge(24, 25, 4, 5, (currency_code), center_normal)
            if Total_Penalties is None:
               ws.write_merge(24, 25, 6, 6, (Total_Penalties), center_normal) 
            else:
                ws.write_merge(24, 25, 6, 6, round(Total_Penalties,2), center_normal)
            
            ws.write_merge(26, 27, 0, 3, "TOTAL ACM IATA BILLED (including ACM issued for Agent Default )",left_normal_bold)
            ws.write_merge(26, 27, 4, 5, (currency_code), center_normal)
            if Total_ACM is None:
               ws.write_merge(26, 27, 6, 6, (Total_ACM), center_normal) 
            else:
                ws.write_merge(26, 27, 6, 6, round(Total_ACM,2), center_normal)
            
            ws.write_merge(28, 29, 0, 3, "TOTAL ADM IATA BILLED ( including ADM issued for Agent Recovery of Default )",left_normal_bold)
            ws.write_merge(28, 29, 4, 5, (currency_code), center_normal)
            if Total_amd is None:
               ws.write_merge(28, 29, 6, 6, (Total_amd), center_normal) 
            else:
                ws.write_merge(28, 29, 6, 6, round(Total_amd,2), center_normal)
            
            ws.write_merge(30, 31, 0, 3, "TOTAL TAX",left_normal_bold)
            ws.write_merge(30, 31, 4, 5, (currency_code), center_normal)
            if Total_tax is None:
               ws.write_merge(30, 31, 6, 6, (Total_tax), center_normal) 
            else:
                ws.write_merge(30, 31, 6, 6, round(Total_tax,2), center_normal)
            
            ws.write_merge(32, 33, 0, 3, "TOTAL BSP SALES EX TAX",left_normal_bold)
            ws.write_merge(32, 33, 4, 5, (currency_code), center_normal)
            if TTL_BSP_SALES_Ex_Tax is None:
               ws.write_merge(32, 33, 6, 6, (TTL_BSP_SALES_Ex_Tax), center_normal) 
            else:
                ws.write_merge(32, 33, 6, 6, round(TTL_BSP_SALES_Ex_Tax,2), center_normal)
            
            ws.write_merge(34, 35, 0, 3, "ACSA  Fee",left_normal_bold)
            ws.write_merge(34, 35, 4, 5, (currency_code), center_normal)
            if ACSA_Fee is None:
               ws.write_merge(34, 35, 6, 6, (ACSA_Fee), center_normal) 
            else:
                ws.write_merge(34, 35, 6, 6, round(ACSA_Fee,2), center_normal)
            
            ws.write_merge(36, 37, 0, 3, "Total",left_normal_border_yellow_center)
            ws.write_merge(36, 37, 4, 5, (currency_code), center_normal_border_yellow_center)
            if ACSA_Fee is None:
               ws.write_merge(36, 37, 6, 6, (ACSA_Fee), center_normal) 
            else:
                ws.write_merge(36, 37, 6, 6, round(ACSA_Fee,2), center_normal_border_yellow_center)
                       
            wb.save(response)
        return response

    def get_all_sundays(self, month, year):
        import calendar

        sundays = []
        cal = calendar.Calendar()

        for day in cal.itermonthdates(year, month):
            if day.weekday() == 6 and day.month == month:
                sundays.append(day)
        return sundays

