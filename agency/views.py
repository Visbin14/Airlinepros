# Standard library imports.
import calendar
import datetime
import string
import os
from collections import OrderedDict

import xlwt
from dateutil import rrule
from django.contrib import messages
from django.contrib.auth.mixins import PermissionRequiredMixin
from django.core.paginator import Paginator
from django.db.models import Count
from django.db.models import Q, Value
from django.db.models import Sum, F,Value as V
from django.db.models.functions import Coalesce
from django.db.models.functions import Concat
from django.http import HttpResponse, HttpResponseRedirect, JsonResponse
from django.shortcuts import render
from django.urls import reverse_lazy
from django.views import View
from django.views.generic import ListView, DetailView, CreateView, UpdateView, DeleteView, TemplateView
from report.models import ReportPeriod
# Local application/library specific imports.
from account.models import User
from agency.forms import AgencyListReferenceForm, AgencyForm, AgencyTypeForm, StateOwnerForm, AgencyCollectionForm
from agency.models import Agency, AgencyType, StatusChange, STATUS_MODES_IATA, AgencyCollection, STATUS_MODES,NewAgents
from agency.tasks import process_agency_list_from_html, process_agency_list_from_csv, process_bulletin, \
    process_agency_list_from_txt
from main.excelstyle import *
from main.models import State, City, Country, Airline
from report.models import Transaction
from main.tasks import send_mail, is_arc


# Related third party imports.
# from bs4 import BeautifulSoup
# from openpyxl import load_workbook

class GetAgencyList(PermissionRequiredMixin, View):
    """Filtered Agency list download as CSV."""

    permission_required = ('agency.view_agency',)


    def get(self, request, *args, **kwargs):
        query = self.request.GET.get('q')
        state = self.request.GET.get('state', None)
        city = self.request.GET.get('city', None)
        owner = self.request.GET.get('owner', None)
        status = self.request.GET.get('status', None)
        alpha = self.request.GET.get('alpha', None)

        country = Country.objects.get(id=request.session.get('country'))

        if country.name != 'United States':
            qs = Agency.csv.select_related('city').values_list("agency_no", "trade_name", "address1", "address2",
                                                           "city__name", "state__name",
                                                           "country__name", "zip_code",
                                                           "email", "tel", "agency_type__name",
                                                           "sales_owner__email", "status_iata")
            header = b'Agency Number,Trade Name,Address1,Address2,City,Provinces,Country,Zip Code,Email,Tel,Agency Type,Sales Owner,Status\n'
        else:
            qs = Agency.csv.select_related('city').values_list("agency_no", "trade_name", "address1", "address2",
                                                               "city__name", "state__name",
                                                               "country__name", "zip_code",
                                                               "email", "tel", "agency_type__name",
                                                               "sales_owner__email", "status")
            header = b'Agency Number,Trade Name,Address1,Address2,City,State,Country,Zip Code,Email,Tel,Agency Type,Sales Owner,Status\n'
        qs = qs.filter(country=self.request.session.get('country'))
        if query:
            qs = qs.filter(
                Q(agency_no__icontains=query) | Q(
                    trade_name__icontains=query) | Q(tel__icontains=query)
            )

        if state:
            qs = qs.filter(state__abrev=state)
        if alpha:
            qs = qs.filter(trade_name__istartswith=alpha)

        if city:
            qs = qs.filter(city=city)
        if owner:
            qs = qs.filter(sales_owner=owner)
        if status and status != 'all':
            qs = qs.filter(status_iata=status)

        agencies = qs.to_csv(header=False)
        agencies = header + agencies
        response = HttpResponse(agencies, content_type='text/csv')
        response['Content-Disposition'] = 'inline; filename=agency_list.csv'
        return response


class AgencyUpload(PermissionRequiredMixin, View):
    """Agency list upload for all country."""

    template_name = 'agency-uploads.html'
    form_class = AgencyListReferenceForm
    permission_required = ('agency.change_agency',)

    def get(self, request, *args, **kwargs):
        form = self.form_class()
        return render(request, self.template_name, {'form': form, 'activate': 'agencies'})

    def post(self, request, *args, **kwargs):
        form = self.form_class(request.POST, request.FILES)
        country = Country.objects.get(id=request.session.get('country'))
        file_error = True
        is_async = False
        to_email = request.user.email
        # to_email = 'basil.jose@fingent.com'
        if form.is_valid():
            obj = form.save()
            file_type = obj.file_type
            extention = obj.file.path.split('.')[-1]

            if file_type == 1:
                if obj.file.size > (20 * 1048576):#20MB
                    is_async = True
                    # Agency List handling goes here
                    if extention.lower() == 'html' and country.name != 'United States':
                        file_error = process_agency_list_from_html.delay(obj.file.path,country.name, is_async=is_async, to_email=to_email)
                    elif extention.lower() == 'txt' and country.name != 'United States':
                        file_error = process_agency_list_from_txt.delay(obj.file.path, country.name,is_async=is_async, to_email=to_email,)
                    elif extention.lower() == 'csv' and country.name == 'United States':
                        file_error = process_agency_list_from_csv.delay(obj.file.path, is_async=is_async, to_email=to_email)
                    else:
                        is_async = False
                        file_error = True
                else:
                    # Agency List handling goes here
                    if extention.lower() == 'html' and country.name != 'United States':
                        file_error = process_agency_list_from_html(obj.file.path,country_name=country.name, is_async=is_async,
                                                                   to_email=to_email)
                    elif extention.lower() == 'txt' and country.name != 'United States':
                        file_error = process_agency_list_from_txt(obj.file.path,country_name=country.name, is_async=is_async,
                                                                  to_email=to_email)
                    elif extention.lower() == 'csv' and country.name == 'United States':
                        file_error = process_agency_list_from_csv(obj.file.path, is_async=is_async,
                                                                  to_email=to_email)
            else:
                # Revokation/Reinstatment handling goes here
                if extention.lower() in ['xlsx', 'xls']:
                    file_error = process_bulletin(obj.file.path, request)

            #async task
            if is_async:
                messages.add_message(self.request, messages.WARNING,
                                     'File upload is taking more time than expected. You will receive an email notification when the upload is completed.')
                return render(request, self.template_name, {'form': form})
            # common error
            if file_error:
                obj.file.delete(False)
                obj.delete()
                messages.add_message(self.request, messages.ERROR,
                                     'Incorrect file format. File has not been imported.')
                return render(request, self.template_name, {'form': form})
        messages.add_message(self.request, messages.SUCCESS,
                             'File uploaded successfully.')
        return render(request, self.template_name, {'form': form})


class AgencyListView(PermissionRequiredMixin, ListView):
    """Agency listing with pagination and ordering."""

    template_name = 'agency-listing.html'
    context_object_name = 'agencies'

    # paginate_by = 20
    permission_required = ('agency.view_agency',)
    
    
    
    def get_queryset(self):
        
        # Transaction.objects.filter(report__report_period__year= 2022, report__country__code = 'AU' ).delete()
        # Transaction.objects.filter(report__report_period__week= 3, report__report_period__year= 2022,report__report_period__month= 8, report__country__code = 'AU' ).delete()
        # Transaction.objects.filter(report__report_period__week= 2, report__report_period__year= 2022,report__report_period__month= 8, report__country__code = 'AU' ).delete()
        # Transaction.objects.filter(report__report_period__week= 1, report__report_period__year= 2022,report__report_period__month= 8, report__country__code = 'AU' ).delete()

        # ReportPeriod.objects.filter(country__code='NZ').delete()

        country = Country.objects.get(id=self.request.session.get('country'))
        order = self.request.GET.get('order_by', 'id')
        qs = Agency.objects.filter(country=self.request.session.get('country')).select_related('city')
        state_abrev = State.objects.filter(country=self.request.session.get('country')).exclude(
            abrev='')

        query = self.request.GET.get('q')
        if country.name != 'United States':
            state = self.request.GET.get('state', state_abrev.first().abrev if state_abrev.exists() else country.code)
        else:
            state = self.request.GET.get('state', state_abrev.first().abrev if state_abrev.exists() else 'US')

        city = self.request.GET.get('city', None)
        owner = self.request.GET.get('owner', None)
        # status = self.request.GET.get('status', None)
        status = self.request.GET.getlist('status', [])
        alpha = self.request.GET.get('alpha', None)

        if query:
            qs = qs.filter(
                Q(agency_no__icontains=query) | Q(
                    trade_name__icontains=query) | Q(tel__icontains=query)
            )
            return qs.order_by(order)


        if city:
            qs = qs.filter(city= city)
            
            return qs.order_by(order)
        # if state:
        #     qs = qs.filter(state__abrev=state)

        if alpha:
            qs = qs.filter(trade_name__istartswith=alpha)
        # 
        # if city:
        #     qs = qs.filter(city=city)
        # 
        # if owner:
        #     # import pdb;pdb.set_trace()
        #     qs = qs.filter(sales_owner=owner)
        # if status and status != 'all':
        if status:
            if country.name != 'United States':
                qs = qs.filter(status_iata__in=status)
            else:
                qs = qs.filter(status__in=status)
        return qs.order_by(order)

    def get_context_data(self, **kwargs):
        context = super(AgencyListView, self).get_context_data(**kwargs)
        context['activate'] = 'agencies'
        context['states'] = State.objects.filter(country=self.request.session.get('country')).exclude(
            abrev='').values_list(
            'abrev', flat=True).order_by('abrev').distinct() or ['CA', ]
        context['alpha'] = list(string.ascii_uppercase)

        country = Country.objects.get(id=self.request.session.get('country'))
        if country.name != 'United States':
            context['status'] = OrderedDict(STATUS_MODES_IATA)
        else:
            context['status'] = OrderedDict(STATUS_MODES)
        context['country_name'] = country.name

        context['sales_owners'] = Agency.objects.exclude(sales_owner__isnull=True).values_list(
            'sales_owner__id', 'sales_owner__email').order_by('sales_owner').distinct('sales_owner')
        context['active_alpha'] = self.request.GET.get(
            'alpha', '')

        if self.request.GET.get('alpha'):
            context['active_state'] = ''
            context['state_filter'] = 'alpha'
            context['cities'] = City.objects.filter(country=self.request.session.get('country')).values_list(
                'id', 'name').order_by('name').distinct('name')
        else:
            context['state_filter'] = 'state'
            context['active_state'] = self.request.GET.get('state', context['states'][0])
            context['cities'] = City.objects.filter(country=self.request.session.get('country'), state__abrev=context['active_state']).values_list(
                'id', 'name').order_by('name').distinct('name')

        context["selected_city"] = self.request.GET.get('city', '')
        context["selected_owner"] = self.request.GET.get('owner', '')
        context["selected_status"] = self.request.GET.getlist('status', [])
        context['selected_sate'] = self.request.GET.get('state', '')
        context['selected_alpha'] = self.request.GET.get('alpha', '')
        context['order_by'] = self.request.GET.get('order_by', 'id')
        context['query'] = self.request.GET.get('q', '')

        return context


class AgencyDetailsView(PermissionRequiredMixin, DetailView):
    """View individual agency details."""

    model = Agency
    template_name = 'agency-details.html'
    context_object_name = 'object'

    permission_required = ('agency.view_agency',)

    def get_context_data(self, **kwargs):
        context = super(AgencyDetailsView, self).get_context_data(**kwargs)
        country = Country.objects.get(id=self.request.session.get('country'))
        context['country'] = country.name
        context['activate'] = 'agencies'
        return context


class AgencySalesDetailsView(PermissionRequiredMixin, View):
    """View individual agency sales details."""

    model = Transaction
    template_name = 'agency-sales-details.html'
    context_object_name = 'object'

    permission_required = ('agency.view_agency',)

    def get(self, request, *args, **kwargs):
        start_month_year = self.request.GET.get('start_month_year', '')
        end_month_year = self.request.GET.get('end_month_year', '')
        start_year = self.request.GET.get('start_year', '')
        end_year = self.request.GET.get('end_year', '')
        airline = self.request.GET.get('airline', '')
        qs = Transaction.objects.select_related('agency').filter(agency=self.kwargs.get('pk'), is_sale=True)
        is_arc_var = is_arc(self.request.session.get('country'))

        if airline:
            qs = qs.filter(report__airline=airline)

        if start_month_year and end_month_year:
            month1 = datetime.datetime.strptime(start_month_year, '%B %Y').month or ''
            year1 = datetime.datetime.strptime(start_month_year, '%B %Y').year or ''

            if month1 and year1:
                start = datetime.datetime(year1, month1, 1)

            month2 = datetime.datetime.strptime(end_month_year, '%B %Y').month or ''
            year2 = datetime.datetime.strptime(end_month_year, '%B %Y').year or ''
            if month2 and year2:
                end = datetime.datetime(year2, month2, calendar.monthrange(year2, month2)[1])
            if start and end:
                # import pdb;pdb.set_trace()
                qs = qs.filter(report__report_period__ped__range=[start, end])

        if start_year and end_year:
            qs = qs.filter(report__report_period__year__range=[start_year, end_year])

        context = dict()
        context['q_string'] = self.request.META['QUERY_STRING']
        context['agency'] = Agency.objects.get(pk=self.kwargs.get('pk'))
        context['submitted'] = self.request.GET.get('submitted', '')
        context['start_month_year'] = self.request.GET.get('start_month_year', '')
        context['end_month_year'] = self.request.GET.get('end_month_year', '')
        context['selected_airline'] = self.request.GET.get('airline', '')
        context['date_filter'] = self.request.GET.get('date_filter', 'month_year')
        context['start_year'] = self.request.GET.get('start_year', '')
        context['end_year'] = self.request.GET.get('end_year', '')
        context['airlines'] = Airline.objects.filter(country=self.request.session.get('country'))
        context['transactions'] = qs
        if start_year and end_year:
            years = [year for year in range(int(start_year), int(end_year) + 1)]
            airline_list = Airline.objects.filter(pk__in=qs.values_list('report__airline', flat=True).distinct(),
                                                  country=self.request.session.get('country'))
            values_list = []
            if not is_arc_var:
                for air in airline_list:
                    row = [air, ]
                    for year in years:
                        row.append(qs.filter(report__report_period__year=year, report__airline=air).aggregate(
                            total=Sum('transaction_amount')).get('total') or 0.0)
                    values_list.append(row)
            else:
                for air in airline_list:
                    row = [air, ]
                    for year in years:
                        total = qs.filter(report__report_period__year=year, report__airline=air).aggregate(total=Coalesce(Sum('transaction_amount'), V(0)) + Coalesce(Sum('pen'), V(0))).get('total') or 0.0
                        row.append(total)
                    values_list.append(row)

            context['values'] = values_list
            context['dates'] = years
            context['airs'] = airline_list
        elif start_month_year and end_month_year:
            dates = list(rrule.rrule(rrule.MONTHLY, dtstart=start, until=end))
            airline_list = Airline.objects.filter(pk__in=qs.values_list('report__airline', flat=True).distinct(),
                                                  country=self.request.session.get('country'))
            values_list = []
            if not  is_arc_var:
                for air in airline_list:
                    row = [air, ]
                    for date in dates:
                        row.append(qs.filter(report__report_period__year=date.year, report__report_period__month=date.month,
                                             report__airline=air).aggregate(
                            total=Sum('transaction_amount')).get('total') or 0.0)
                    values_list.append(row)
            else:
                for air in airline_list:
                    row = [air, ]
                    for date in dates:
                        total = qs.filter(report__report_period__year=date.year, report__report_period__month=date.month,
                                             report__airline=air).aggregate(total=Coalesce(Sum('transaction_amount'), V(0)) + Coalesce(Sum('pen'), V(0))).get('total') or 0.0
                        row.append(total)
                    values_list.append(row)
            context['dates'] = dates
            context['values'] = values_list
            context['airs'] = airline_list

        return render(self.request, self.template_name, context)


class AgencySalesExelView(View):
    """Filtered Sales Report for agency download as exel."""

    # permission_required = ('agency.view_agency',)

    def get(self, request, *args, **kwargs):
        start_month_year = self.request.GET.get('start_month_year', '')
        end_month_year = self.request.GET.get('end_month_year', '')
        start_year = self.request.GET.get('start_year', '')
        end_year = self.request.GET.get('end_year', '')
        airline = self.request.GET.get('airline', '')
        qs = Transaction.objects.select_related('agency').filter(agency=self.kwargs.get('pk'), is_sale=True)
        is_arc_var = is_arc(self.request.session.get('country'))
        if airline:
            qs = qs.filter(report__airline=airline)

        response = HttpResponse(content_type='application/vnd.ms-excel')

        trade_name = Agency.objects.get(pk=self.kwargs.get('pk')).trade_name
        if not trade_name:
            trade_name = ''

        file_name = trade_name + "- Sales Details" + ".xls"
        response['Content-Disposition'] = 'inline; filename=' + file_name

        wb = xlwt.Workbook(encoding='utf-8')
        ws = wb.add_sheet('sales details')

        # Sheet header, first row
        row_num = 0

        font_style = xlwt.XFStyle()
        font_style.font.bold = True

        columns = ["Dates :-", ]
        ws.row(row_num).height = 20 * 20
        ws.write_merge(row_num, 0, 0, 6, trade_name.upper(), bold_center)
        row_num = row_num + 1
        ws.row(row_num).height = 20 * 20
        ws.write_merge(row_num, 1, 0, 6, "AGENCY SALES DETAILS REPORT", bold_center)
        row_num = row_num + 1
        ws.row(row_num).height = 20 * 20


        if start_year and end_year:
            ws.write_merge(row_num, 2, 0, 6, "BETWEEN " + start_year.upper() + " AND " + end_year.upper(), bold_center)
            row_num = row_num + 1
            ws.row(row_num).height_mismatch = True

            qs = qs.filter(report__report_period__year__range=[start_year, end_year])
            years = [year for year in range(int(start_year), int(end_year) + 1)]
            columns = columns + years
            airline_list = Airline.objects.filter(pk__in=qs.values_list('report__airline', flat=True).distinct(),
                                                  country=self.request.session.get('country'))
            values_list = []
            if not is_arc_var:
                for air in airline_list:
                    row = [air.name, ]
                    for year in years:
                        row.append(qs.filter(report__report_period__year=year, report__airline=air).aggregate(
                            total=Sum('transaction_amount')).get('total') or 0.0)
                    values_list.append(row)
            else:
                for air in airline_list:
                    row = [air.name, ]
                    for year in years:
                        total = qs.filter(report__report_period__year=year, report__airline=air).aggregate(
                            total=Coalesce(Sum('transaction_amount'), V(0)) + Coalesce(Sum('pen'), V(0))).get('total') or 0.0
                        row.append(total)
                    values_list.append(row)

        elif start_month_year and end_month_year:
            ws.write_merge(row_num, 2, 0, 6, "BETWEEN " + start_month_year.upper() + " AND " + end_month_year.upper(), bold_center)
            row_num = row_num + 1
            ws.row(row_num).height_mismatch = True
            month1 = datetime.datetime.strptime(start_month_year, '%B %Y').month or ''
            year1 = datetime.datetime.strptime(start_month_year, '%B %Y').year or ''

            if month1 and year1:
                start = datetime.datetime(year1, month1, 1)

            month2 = datetime.datetime.strptime(end_month_year, '%B %Y').month or ''
            year2 = datetime.datetime.strptime(end_month_year, '%B %Y').year or ''
            if month2 and year2:
                end = datetime.datetime(year2, month2, calendar.monthrange(year2, month2)[1])
            if start and end:
                # import pdb;pdb.set_trace()
                qs = qs.filter(report__report_period__ped__range=[start, end])

            dates = list(rrule.rrule(rrule.MONTHLY, dtstart=start, until=end))
            columns = columns + [str(date.month) + "-" + str(date.year) for date in dates]
            airline_list = Airline.objects.filter(pk__in=qs.values_list('report__airline', flat=True).distinct(),
                                                  country=self.request.session.get('country'))
            values_list = []
            if not is_arc_var:
                for air in airline_list:
                    row = [air.name, ]
                    for date in dates:
                        row.append(
                            qs.filter(report__report_period__year=date.year, report__report_period__month=date.month,
                                      report__airline=air).aggregate(
                                total=Sum('transaction_amount')).get('total') or 0.0)
                    values_list.append(row)
            else:
                for air in airline_list:
                    row = [air.name, ]
                    for date in dates:
                        total = qs.filter(report__report_period__year=date.year,
                                          report__report_period__month=date.month,
                                          report__airline=air).aggregate(
                            total=Coalesce(Sum('transaction_amount'), V(0)) + Coalesce(Sum('pen'), V(0))).get('total') or 0.0
                        row.append(total)
                    values_list.append(row)
        for i, col in enumerate(columns):
            ws.write(row_num, i, col, yellow_background_header)

        for row in values_list:
            row_num = row_num + 1
            for i, item in enumerate(row):
                ws.write(row_num, i, item)

        # taxes = set(qs.values_list("taxes__type", flat=True))
        #
        # peds = qs.values_list('report__report_period__ped', flat=True).order_by('report__report_period__ped').distinct()
        #
        # for tx_type in taxes:
        #     if tx_type:
        #         columns.append("Tax "+tx_type)
        # # mapping = dict((value, key) for (key, value) in enumerate(columns))
        #
        # for ped in peds:
        #
        #     ws.write(row_num, 0, "PED: " + str(ped), font_style)
        #     row_num = row_num + 1
        #
        #     ws.write(row_num, 0, "")
        #     row_num = row_num + 1
        #
        #     for col_num in range(len(columns)):
        #         ws.write(row_num, col_num, columns[col_num], font_style)
        #
        #     for item in qs.filter(report__report_period__ped=ped):
        #         row_num = row_num+1
        #         tax_start = len(default_columns)
        #         ws.write(row_num, 0, item.card_type or '')
        #         ws.write(row_num, 1, item.agency.agency_no or '')
        #         ws.write(row_num, 2, item.issue_date or '')
        #         ws.write(row_num, 3, item.ticket_no or '')
        #         ws.write(row_num, 4, item.transaction_amount or '')
        #         ws.write(row_num, 5, item.cobl_amount or '')
        #         ws.write(row_num, 6, item.std_comm_amount or '')
        #         ws.write(row_num, 7, item.std_comm_rate or '')
        #         ws.write(row_num, 8, item.transaction_type or '')
        #         ws.write(row_num, 9, item.pen or '')
        #
        #         for tax in taxes:
        #             if tax:
        #                 tax_amts = [tx for tx in item.taxes_set.all() if tx.type == tax]
        #                 if len(tax_amts)>0:
        #                     tax_amt = tax_amts[0].amount
        #                     ws.write(row_num, tax_start, tax_amt)
        #
        #                 tax_start = tax_start + 1
        #     row_num = row_num + 1
        #     ws.write(row_num, 0, "")
        #     row_num = row_num + 1

        # Sheet body, remaining rows
        font_style = xlwt.XFStyle()

        wb.save(response)
        return response


class AgencyUpdateView(PermissionRequiredMixin, UpdateView):
    """ Edit agency details."""

    form_class = AgencyForm
    model = Agency
    template_name = 'agency-edit.html'

    context_object_name = 'object'
    permission_required = ('agency.change_agency',)

    def form_valid(self, form):
        messages.add_message(self.request, messages.SUCCESS,
                             'Agency details updated successfully.')
        return super(AgencyUpdateView, self).form_valid(form)

    def get_context_data(self, **kwargs):
        context = super(AgencyUpdateView, self).get_context_data(**kwargs)
        context['activate'] = 'agencies'
        return context

    def get_form_kwargs(self):
        kwargs = super(AgencyUpdateView, self).get_form_kwargs()
        kwargs['agency_types'] = AgencyType.objects.filter(country=self.request.session.get('country'))
        kwargs['cities'] = City.objects.filter(country=self.request.session.get('country'))
        kwargs['states'] = State.objects.filter(country=self.request.session.get('country'))
        return kwargs


class AgencyTypeView(PermissionRequiredMixin, ListView):
    """AgencyType listing with pagination and ordering."""

    template_name = 'agency-types.html'
    context_object_name = 'agency_types'

    permission_required = ('agency.view_agencytype',)

    def get_queryset(self):
        order = self.request.GET.get('order_by', 'id')
        qs = AgencyType.objects.filter(country=self.request.session.get('country')).annotate(
            assigned_number=Count('agencies'))
        return qs.order_by(order)

    def get_context_data(self, **kwargs):
        context = super(AgencyTypeView, self).get_context_data(**kwargs)
        context['activate'] = 'agencies'
        context['order_by'] = self.request.GET.get('order_by', 'id')
        return context


class AgencyTypeDetailsView(PermissionRequiredMixin, DetailView):
    """AgencyType details view."""

    model = AgencyType
    template_name = 'agency-type-details.html'
    context_object_name = 'object'

    page_size = 20
    permission_required = ('agency.view_agencytype',)

    def get_context_data(self, **kwargs):
        context = super(AgencyTypeDetailsView, self).get_context_data(**kwargs)
        city = self.request.GET.get('city', None)
        state = self.request.GET.get('state', None)
        page = self.request.GET.get('page', 1)
        order = self.request.GET.get('order_by', 'id')

        qs = Agency.objects.select_related('city', 'state').filter(agency_type=self.object)

        if state:
            qs = qs.filter(state=state)

        if city:
            qs = qs.filter(city=city)

        qs = qs.order_by(order)
        paginator = Paginator(qs, self.page_size)
        qs = paginator.get_page(page)

        context['activate'] = 'agencies'
        context['page'] = int(page)
        context['paginator'] = paginator
        context['states'] = State.objects.exclude(abrev='').filter(
            country=self.request.session.get('country')).values_list(
            'id', 'name').order_by('abrev')
        context['cities'] = City.objects.values_list(
            'name', 'id').filter(country=self.request.session.get('country')).order_by('name')
        context['order_by'] = self.request.GET.get('order_by')
        context['agencies'] = qs
        context["selected_city"] = self.request.GET.get('city', '')
        context["selected_state"] = self.request.GET.get('state', '')
        return context


class AgencyTypeCreate(PermissionRequiredMixin, CreateView):
    """Create new AgencyType"""

    model = AgencyType
    form_class = AgencyTypeForm
    template_name = 'agency-type-add.html'
    success_url = reverse_lazy('agency_types')
    permission_required = ('agency.change_agencytype',)

    def form_valid(self, form):
        country = self.request.session.get('country')

        cleaned_data = form.cleaned_data
        name = cleaned_data.get('name')
        if name and AgencyType.objects.filter(name__iexact=name, country=country).exists():
            form.add_error('name', 'AgencyType with that name already exists.')
            return super(AgencyTypeCreate, self).form_invalid(form)

        self.object = form.save(commit=False)
        self.object.country = Country.objects.get(id=country)
        self.object.save()
        agency_nos = form.cleaned_data.get('agencies', '')
        message = 'Agency Type created successfully.'
        message_type = messages.SUCCESS
        if agency_nos:
            agency_nos = list(filter(None, agency_nos.split(',')))
            agencies = Agency.objects.filter(
                agency_no__in=agency_nos, country=country).update(agency_type=self.object)
        not_found = []
        for age in agency_nos:
            if not Agency.objects.filter(agency_no=age).exists():
                not_found.append(age)
        if not_found:
            message = "Agency Type created! agency number not found : " + ",".join(not_found)
            message_type = messages.WARNING
        messages.add_message(self.request, message_type, message)
        return HttpResponseRedirect(self.get_success_url())


class AgencyTypeUpdateView(PermissionRequiredMixin, UpdateView):
    """Update AgencyType details."""

    model = AgencyType
    form_class = AgencyTypeForm
    template_name = 'agency-type-edit.html'

    context_object_name = 'object'
    permission_required = ('agency.change_agencytype',)

    def form_valid(self, form):
        country = self.request.session.get('country')

        cleaned_data = form.cleaned_data
        changed_data = form.changed_data
        if 'name' in changed_data:
            if AgencyType.objects.filter(name__iexact=cleaned_data.get('name'), country=country).exists():
                form.add_error('name', 'AgencyType with that name already exists.')
                return super(AgencyTypeUpdateView, self).form_invalid(form)

        self.object = form.save()
        agency_nos = form.cleaned_data.get('agencies', '')
        message = 'Agency Type updated successfully.'
        message_type = messages.SUCCESS
        if agency_nos:
            agency_nos = list(filter(None, agency_nos.split(',')))
            agencies = Agency.objects.filter(
                agency_no__in=agency_nos).update(agency_type=self.object)
            not_found = []
            for age in agency_nos:
                if not Agency.objects.filter(agency_no=age).exists():
                    not_found.append(age)
            if not_found:
                message = "The following agency number were not found : " + ",".join(not_found)
                message_type = messages.WARNING
        messages.add_message(self.request, message_type, message)
        return HttpResponseRedirect(self.get_success_url())

    def get_context_data(self, **kwargs):
        context = super(AgencyTypeUpdateView, self).get_context_data(**kwargs)
        context['activate'] = 'agencies'
        return context


class AgencyTypeDelete(PermissionRequiredMixin, DeleteView):
    """Delete AgencyType"""

    model = AgencyType
    success_url = reverse_lazy('agency_types')
    permission_required = ('agency.change_agencytype',)

    def delete(self, *args, **kwargs):
        self.get_object().delete()
        messages.add_message(self.request, messages.SUCCESS,
                             'Agency Type deleted successfully.')
        return HttpResponseRedirect(self.success_url)


class AgencyTypeRemoveAgency(PermissionRequiredMixin, View):
    """Remove an agency from AgencyType"""

    permission_required = ('agency.change_agencytype',)

    def post(self, *args, **kwargs):
        agency_id = self.request.POST.get('agency_id')
        try:
            agency = Agency.objects.get(id=agency_id)
            agency.agency_type = None
            agency.save()
            response = JsonResponse({'message': 'Agency successfully removed'})
        except Exception as e:
            response = JsonResponse({'error': 'Agency not found'}, status=400)
        return response


class AgencyCollectionView(PermissionRequiredMixin, ListView):
    """Agency Collection listing with pagination and ordering."""

    template_name = 'agency-collections.html'
    context_object_name = 'agency_collections'
    permission_required = ('agency.view_agencycollection',)

    def get_queryset(self):
        order = self.request.GET.get('order_by', 'id')
        qs = AgencyCollection.objects.filter(country=self.request.session.get('country')).annotate(
            assigned_number=Count('collection_agencies'))
        return qs.order_by(order)

    def get_context_data(self, **kwargs):
        context = super(AgencyCollectionView, self).get_context_data(**kwargs)
        context['activate'] = 'agencies'
        context['order_by'] = self.request.GET.get('order_by', 'id')
        return context


class AgencyCollectionDetailsView(PermissionRequiredMixin, DetailView):
    """AgencyCollection details view."""

    model = AgencyCollection
    template_name = 'agency-collection-details.html'
    context_object_name = 'object'
    permission_required = ('agency.view_agencycollection',)
    page_size = 20

    def get_context_data(self, **kwargs):
        context = super(AgencyCollectionDetailsView, self).get_context_data(**kwargs)
        city = self.request.GET.get('city', None)
        state = self.request.GET.get('state', None)
        page = self.request.GET.get('page', 1)
        order = self.request.GET.get('order_by', 'id')

        qs = Agency.objects.select_related('city', 'state').filter(agency_collection=self.object)

        if state:
            qs = qs.filter(state=state)

        if city:
            qs = qs.filter(city=city)

        qs = qs.order_by(order)
        paginator = Paginator(qs, self.page_size)
        qs = paginator.get_page(page)

        context['activate'] = 'agencies'
        context['page'] = int(page)
        context['paginator'] = paginator
        context['states'] = State.objects.exclude(abrev='').filter(
            country=self.request.session.get('country')).values_list(
            'id', 'name').order_by('abrev')
        context['cities'] = City.objects.values_list(
            'name', 'id').filter(country=self.request.session.get('country')).order_by('name')
        context['order_by'] = self.request.GET.get('order_by')
        context['agencies'] = qs
        context["selected_city"] = self.request.GET.get('city', '')
        context["selected_state"] = self.request.GET.get('state', '')
        return context


class AgencyCollectionCreate(PermissionRequiredMixin, CreateView):
    """Create new AgencyCollection"""

    model = AgencyCollection
    form_class = AgencyCollectionForm
    template_name = 'agency-collection-add.html'
    success_url = reverse_lazy('agency_collections')
    permission_required = ('agency.change_agencycollection',)

    def form_valid(self, form):
        country = self.request.session.get('country')

        cleaned_data = form.cleaned_data
        name = cleaned_data.get('name')
        if name and AgencyCollection.objects.filter(name__iexact=name, country=country).exists():
            form.add_error('name', 'Agency Collection with that name already exists.')
            return super(AgencyCollectionCreate, self).form_invalid(form)

        self.object = form.save(commit=False)
        self.object.country = Country.objects.get(id=country)
        self.object.save()
        agency_nos = form.cleaned_data.get('agencies', '')
        message = 'Agency Collection created successfully.'
        message_type = messages.SUCCESS
        if agency_nos:
            agency_nos = list(filter(None, agency_nos.split(',')))
            agencies = Agency.objects.filter(
                agency_no__in=agency_nos, country=country).update(agency_collection=self.object)
        not_found = []
        for age in agency_nos:
            if not Agency.objects.filter(agency_no=age).exists():
                not_found.append(age)
        if not_found:
            message = "Agency Collection created! agency number not found : " + ",".join(not_found)
            message_type = messages.WARNING
        messages.add_message(self.request, message_type, message)
        return HttpResponseRedirect(self.get_success_url())


class AgencyCollectionUpdateView(PermissionRequiredMixin, UpdateView):
    """Update AgencyCollection details."""

    model = AgencyCollection
    form_class = AgencyCollectionForm
    template_name = 'agency-collection-edit.html'

    context_object_name = 'object'
    permission_required = ('agency.change_agencycollection',)

    def form_valid(self, form):
        country = self.request.session.get('country')

        cleaned_data = form.cleaned_data
        changed_data = form.changed_data
        if 'name' in changed_data:
            if AgencyCollection.objects.filter(name__iexact=cleaned_data.get('name'), country=country).exists():
                form.add_error('name', 'Agency Collection with that name already exists.')
                return super(AgencyCollectionUpdateView, self).form_invalid(form)

        self.object = form.save()
        agency_nos = form.cleaned_data.get('agencies', '')
        message = 'Agency Collection updated successfully.'
        message_type = messages.SUCCESS
        if agency_nos:
            agency_nos = list(filter(None, agency_nos.split(',')))
            agencies = Agency.objects.filter(
                agency_no__in=agency_nos, country=country).update(agency_collection=self.object)
            not_found = []
            for age in agency_nos:
                if not Agency.objects.filter(agency_no=age).exists():
                    not_found.append(age)
            if not_found:
                message = "The following agency numbers were not found : " + ",".join(not_found)
                message_type = messages.WARNING
        messages.add_message(self.request, message_type, message)
        return HttpResponseRedirect(self.get_success_url())

    def get_context_data(self, **kwargs):
        context = super(AgencyCollectionUpdateView, self).get_context_data(**kwargs)
        context['activate'] = 'agencies'
        return context


class AgencyCollectionDelete(PermissionRequiredMixin, DeleteView):
    """Delete AgencyCollection"""

    model = AgencyCollection
    success_url = reverse_lazy('agency_collections')
    permission_required = ('agency.change_agencycollection',)

    def delete(self, *args, **kwargs):
        self.get_object().delete()
        messages.add_message(self.request, messages.SUCCESS,
                             'Agency Collection deleted successfully.')
        return HttpResponseRedirect(self.success_url)


class AgencyCollectionRemoveAgency(PermissionRequiredMixin, View):
    """Remove an agency from AgencyCollection"""

    permission_required = ('agency.change_agencycollection',)

    def post(self, *args, **kwargs):
        agency_id = self.request.POST.get('agency_id')
        try:
            agency = Agency.objects.get(id=agency_id)
            agency.agency_collection = None
            agency.save()
            response = JsonResponse({'message': 'Agency successfully removed'})
        except Exception as e:
            response = JsonResponse({'error': 'Agency not found'}, status=400)
        return response


class StatusHistoryView(PermissionRequiredMixin, TemplateView):
    """StatusHistory listing with pagination and ordering."""

    template_name = 'agency-status-history.html'

    paginate_by = 20
    permission_required = ('agency.view_agency',)

    def get_context_data(self, **kwargs):
        context = super(StatusHistoryView, self).get_context_data(**kwargs)
        context['activate'] = 'agencies'
        country = Country.objects.get(id=self.request.session.get('country'))
        if country.name != 'United States':
            context['status'] = OrderedDict(STATUS_MODES_IATA)
        else:
            context['status'] = OrderedDict(STATUS_MODES)

        context['order_by'] = self.request.GET.get('order_by', 'id')
        context['agency'] = Agency.objects.get(pk=kwargs.get('pk'))

        # Remove current status from dropdown
        # del(context['status'][context['agency'].status])

        context['histories'] = StatusChange.objects.filter(
            agency=kwargs.get('pk')).order_by(context['order_by'])
        return context

    def post(self, request, *args, **kwargs):
        pk = kwargs.get('pk')
        reason = self.request.POST.get('reason', None)
        status = self.request.POST.get('status', None)
        agency = Agency.objects.get(pk=pk)
        country = Country.objects.get(id=self.request.session.get('country'))
        if country.name != 'United States':
            old_status = agency.get_status_iata_display()
            agency.status_iata = status
            agency.save()
            new_status = agency.get_status_iata_display()
        else:
            old_status = agency.get_status_display()
            agency.status = status
            agency.save()
            new_status = agency.get_status_display()
        StatusChange.objects.create(
            old_status=old_status, new_status=new_status, reason=reason, agency=agency)
        if (new_status == 'Default Information' and country.name != 'United States') or (new_status == 'Revoked' and country.name == 'United States'):
            if new_status == 'Default Information':
                agency_status = 'defaulted'
            else:
                agency_status = 'revoked'
            context = {
                'agency_no': agency.agency_no,
                'agency_name': agency.trade_name,
                'agency_status': agency_status
            }
            to = agency.sales_owner.email if agency.sales_owner else ''
            send_mail("Agency status changed.", "email/status-default-email.html", context,
                      [to],
                      from_email='Assda@assda.com')
        messages.add_message(self.request, messages.SUCCESS,
                             'Status changed successfully.')
        return HttpResponseRedirect(reverse_lazy("status_history", kwargs={'pk': pk}))


class StateOwnersListView(PermissionRequiredMixin, ListView):
    """StatesSalesOwners listing with pagination and ordering."""

    model = User
    template_name = 'state-owner-listing.html'
    context_object_name = 'owners'

    paginate_by = 20
    permission_required = ('main.view_stateowners',)

    def get_queryset(self):
        order = self.request.GET.get('order_by', 'name')
        qs = User.objects.annotate(name=Concat(
            'first_name', Value(' '), 'last_name'))
        query = self.request.GET.get('q')
        if query:
            qs = qs.filter(
                Q(name__icontains=query) | Q(state__name__icontains=query)
            )
        return qs.order_by(order)

    def get_context_data(self, **kwargs):
        context = super(StateOwnersListView, self).get_context_data(**kwargs)
        ownerState = {}
        try:
            qry = State.objects.filter(owner__in=context['owners'].values_list('id'), country__id=self.request.session.get('country'))
            if qry:
                for q in qry:
                    if q.owner_id not in ownerState:
                        ownerState[q.owner_id] = []

                    ownerState[q.owner_id].append(q.name)
        except Exception as e:
            pass

        context['activate'] = 'agencies'
        context['ownerState'] = ownerState
        context['order_by'] = self.request.GET.get('order_by', 'id')
        context['query'] = self.request.GET.get('q', '')
        return context


class StateOwnerUpdateView(PermissionRequiredMixin, UpdateView):
    """Edit StateOwner details."""

    form_class = StateOwnerForm
    model = User
    template_name = 'edit-state-owner.html'

    context_object_name = 'object'
    success_url = reverse_lazy('state_owners')
    permission_required = ('main.change_stateowners',)

    def get_form_kwargs(self):
        kwargs = super(StateOwnerUpdateView, self).get_form_kwargs()
        kwargs['user'] = self.kwargs['pk']  # pass the 'user' in kwargs
        kwargs['request'] = self.request  # pass the 'request' in kwargs
        return kwargs

    def form_valid(self, form):
        self.object = form.save(commit=False)
        State.objects.filter(owner=self.object).update(owner='')
        form.cleaned_data['states'].update(owner=self.object)
        messages.add_message(self.request, messages.SUCCESS,
                             'Updated successfully.')
        return HttpResponseRedirect(self.get_success_url())

    def test_func(self):
        return True

    def get_context_data(self, **kwargs):
        context = super(StateOwnerUpdateView, self).get_context_data(**kwargs)
        context['activate'] = 'agencies'
        return context


class AgencyCollectionReportView(PermissionRequiredMixin, View):
    """View agency collection sales details."""

    model = Transaction
    template_name = 'agency-collection-report.html'
    context_object_name = 'object'

    permission_required = ('report.view_agency_collection_report',)

    def get(self, request, *args, **kwargs):
        start_month_year = self.request.GET.get('start_month_year', '')
        end_month_year = self.request.GET.get('end_month_year', '')
        airline = self.request.GET.get('airline', '')

        qs = Transaction.objects.filter(agency__agency_collection=self.kwargs.get('pk'), is_sale=True)
        context = dict()
        context['q_string'] = self.request.META['QUERY_STRING']
        context['submitted'] = self.request.GET.get('submitted', '')
        context['start_month_year'] = self.request.GET.get('start_month_year', '')
        context['end_month_year'] = self.request.GET.get('end_month_year', '')
        context['selected_airline'] = self.request.GET.get('airline', '')
        context['organize_by'] = self.request.GET.get('organize_by', 'year_range')
        context['airlines'] = Airline.objects.filter(country=self.request.session.get('country'))
        context['collection'] = AgencyCollection.objects.get(pk=self.kwargs.get('pk'))
        if airline:
            qs = qs.filter(report__airline=airline)

        if start_month_year and end_month_year:
            month1 = datetime.datetime.strptime(start_month_year, '%B %Y').month or ''
            year1 = datetime.datetime.strptime(start_month_year, '%B %Y').year or ''

            if month1 and year1:
                start = datetime.datetime(year1, month1, 1)

            month2 = datetime.datetime.strptime(end_month_year, '%B %Y').month or ''
            year2 = datetime.datetime.strptime(end_month_year, '%B %Y').year or ''
            if month2 and year2:
                end = datetime.datetime(year2, month2, calendar.monthrange(year2, month2)[1])
            if start and end:
                qs = qs.filter(report__report_period__ped__range=[start, end])

            context['transactions'] = qs
            if context['organize_by'] == 'year_range':

                years = [year for year in range(int(year1), int(year2) + 1)]
                annotations = dict()
                for y in years:
                    annotation_name = '{}'.format(y)
                    annotations[annotation_name] = Sum('fare_amount',
                                                       filter=Q(report__report_period__year=y))
                values_list = qs.values('agency').order_by().distinct().annotate(trade_name=F('agency__trade_name'),
                                                                      agency_no=F('agency__agency_no'),
                                                                      city=F('agency__city__name'),state=F('agency__state__name'),agency_type=F('agency__agency_type__name'), **annotations)
                context['values'] = values_list
                context['dates'] = [str(year) for year in years]
            else:
                dates = list(rrule.rrule(rrule.MONTHLY, dtstart=start, until=end))
                annotations = dict()
                for date in dates:
                    annotation_name = date.strftime('%m-%Y')
                    annotations[annotation_name] = Sum('fare_amount',
                                                       filter=Q(report__report_period__year=date.year, report__report_period__month=date.month))
                values_list = qs.values('agency').order_by().distinct().annotate(trade_name=F('agency__trade_name'),
                                                                      agency_no=F('agency__agency_no'),
                                                                      city=F('agency__city__name'),state=F('agency__state__name'),agency_type=F('agency__agency_type__name'), **annotations)

                context['dates'] = [date.strftime('%m-%Y') for date in dates]
                context['values'] = values_list

        return render(self.request, self.template_name, context)


class AgencyCollectionReportDownloadView(PermissionRequiredMixin, View):
    """Download agency collection sales details."""

    permission_required = ('report.download_agency_collection_report',)

    def get(self, request, *args, **kwargs):
        start_month_year = self.request.GET.get('start_month_year', '')
        end_month_year = self.request.GET.get('end_month_year', '')
        airline = self.request.GET.get('airline', '')
        organize_by = self.request.GET.get('organize_by', '')

        qs = Transaction.objects.filter(agency__agency_collection=self.kwargs.get('pk'), is_sale=True)
        context = dict()
        context['collection'] = AgencyCollection.objects.get(pk=self.kwargs.get('pk'))
        if airline:
            qs = qs.filter(report__airline=airline)

        if start_month_year and end_month_year:
            month1 = datetime.datetime.strptime(start_month_year, '%B %Y').month or ''
            year1 = datetime.datetime.strptime(start_month_year, '%B %Y').year or ''

            if month1 and year1:
                start = datetime.datetime(year1, month1, 1)

            month2 = datetime.datetime.strptime(end_month_year, '%B %Y').month or ''
            year2 = datetime.datetime.strptime(end_month_year, '%B %Y').year or ''
            if month2 and year2:
                end = datetime.datetime(year2, month2, calendar.monthrange(year2, month2)[1])
            if start and end:
                qs = qs.filter(report__report_period__ped__range=[start, end])

            if organize_by == 'year_range':

                years = [year for year in range(int(year1), int(year2) + 1)]
                annotations = dict()
                for y in years:
                    annotation_name = '{}'.format(y)
                    annotations[annotation_name] = Sum('fare_amount',
                                                       filter=Q(report__report_period__year=y))
                values_list = qs.values('agency').order_by().distinct().annotate(trade_name=F('agency__trade_name'),
                                                                      agency_no=F('agency__agency_no'),
                                                                      city=F('agency__city__name'),state=F('agency__state__name'),agency_type=F('agency__agency_type__name'), **annotations)
                context['values'] = values_list
                context['dates'] = [str(year) for year in years]
            else:
                dates = list(rrule.rrule(rrule.MONTHLY, dtstart=start, until=end))
                annotations = dict()
                for date in dates:
                    annotation_name = date.strftime('%m-%Y')
                    annotations[annotation_name] = Sum('fare_amount',
                                                       filter=Q(report__report_period__year=date.year, report__report_period__month=date.month))
                values_list = qs.values('agency').order_by().distinct().annotate(trade_name=F('agency__trade_name'),
                                                                      agency_no=F('agency__agency_no'),
                                                                      city=F('agency__city__name'),state=F('agency__state__name'),agency_type=F('agency__agency_type__name'), **annotations)

                context['dates'] = [date.strftime('%m-%Y') for date in dates]
                context['values'] = values_list

        response = HttpResponse(content_type='application/vnd.ms-excel')
        wb = xlwt.Workbook(encoding='utf-8')
        ws = FitSheetWrapper(wb.add_sheet('Agency Collection Report'))
        row_num = 0
        if airline:
            airline_obj = Airline.objects.filter(id=airline).first()
            if airline_obj:
                ws.write_merge(row_num, 2, 0, 10,
                               "Sales report of " + airline_obj.name + " For " + context['collection'].name + ' ' + start_month_year + ' to ' + end_month_year, bold_center)
        else:
            ws.write_merge(row_num, 2, 0, 10,
                           "Sales report For " + context[
                               'collection'].name + ' ' + start_month_year + ' to ' + end_month_year,
                           bold_center)
        row_num = row_num + 3
        file_name = "collection_report.xls"
        response['Content-Disposition'] = 'inline; filename=' + file_name

        ws.write(row_num, 0, "Agency trade name", yellow_background_header)
        ws.write(row_num, 1, "Agency number ", yellow_background_header)
        ws.write(row_num, 2, "City", yellow_background_header)
        ws.write(row_num, 3, "State", yellow_background_header)
        ws.write(row_num, 4, "Agency type", yellow_background_header)
        for i, date in enumerate(context['dates']):
            ws.write(row_num, i + 5, date, yellow_background_header)

        row_num = row_num + 1

        for value in context['values']:
            ws.write(row_num, 0, value.get('trade_name'))
            ws.write(row_num, 1, value.get('agency_no'))
            ws.write(row_num, 2, value.get('city'))
            ws.write(row_num, 3, value.get('state'))
            ws.write(row_num, 4, value.get('agency_type'))
            for i, date in enumerate(context['dates']):
                ws.write(row_num, i+5, value.get(date) or '0.0')
            row_num = row_num + 1

        wb.save(response)
        return response


class NewAgentReportView(ListView):
    model = NewAgents
    template_name = "new-agent-list.html"
    context_object_name = 'agent_list'
    def get_queryset(self):
        order = self.request.GET.get('order_by', '-created_on')
        airline = self.request.GET.get('airline', '')
        ped = self.request.GET.get('ped', '')
        if airline and ped:
            qs = NewAgents.objects.filter(airline__id=airline,ped=ped)
            query = self.request.GET.get('q')
            if query:
                qs = qs.filter(
                    Q(agency__trade_name__icontains=query) | Q(agency__agency_no__icontains=query)
                )
            return qs.order_by(order)

    def get_context_data(self, **kwargs):
        context = super(NewAgentReportView, self).get_context_data(**kwargs)
        context['order_by'] = self.request.GET.get('order_by', '-created_on')
        context['selected_airline'] = self.request.GET.get('airline', '')
        newagents = NewAgents.objects.filter(airline__country=self.request.session.get('country'))
        ped_list = list(newagents.distinct('ped').values_list('ped',flat=True))
        airlines = list(newagents.distinct(
            'airline').values_list('airline', flat=True))
        context['ped_list'] = ped_list
        context['selected_ped'] = self.request.GET.get('ped')
        context['airlines'] = [Airline.objects.get(id=airline) for airline in airlines]
        return context

class NewAgentReportDownloadView(View):

    def get(self,request):
        response = HttpResponse(content_type='application/vnd.ms-excel')
        airline = request.GET.get('airline', '')
        ped = self.request.GET.get('ped', '')
        if airline and ped:
            airline_obj = Airline.objects.filter(pk=airline).first()
            file_name = "New Agent Report for {} - PED {}.xls".format(airline_obj.name,ped)
            response['Content-Disposition'] = 'inline; filename=' + file_name
            wb = xlwt.Workbook(encoding='utf-8')
            ws = FitSheetWrapper(wb.add_sheet('sheet1'))
            row_num = 0
            font_style = xlwt.XFStyle()
            font_style.font.bold = True
            style_string = "font: bold on; pattern: pattern solid, fore_color yellow; borders: top_color black, bottom_color black, right_color black, left_color black,left thin, right thin, top thin, bottom thin;"
            yellow_style = xlwt.easyxf(style_string)
            ws.row(row_num).height = 20 * 20
            ws.write_merge(row_num, 0, 0, 6, airline_obj.name.upper(), bold_center)
            columns = ["Agent no", "Agent Name", "City", "Telephone", "Email"]
            row_num += 1
            column_width = 5000
            for index in range(len(columns)):
                ws.write(row_num,index, columns[index], yellow_style)
            for agent in NewAgents.objects.filter(airline__id=airline_obj.id,ped=ped).order_by('-created_on'):
                row_num = row_num + 1
                ws.write(row_num, 0, agent.agency.agency_no)
                ws.write(row_num, 1, agent.agency.trade_name)
                ws.write(row_num, 2, agent.agency.city.name if agent.agency.city else "")
                ws.write(row_num, 3, agent.agency.tel)
                ws.write(row_num, 4, agent.agency.email)
                # ws.col(0).width = column_width
                # ws.col(1).width = column_width
                # ws.col(2).width = column_width
                # ws.col(3).width = column_width
                # ws.col(4).width = column_width
            wb.save(response)
        return response