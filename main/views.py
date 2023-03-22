import json
import os
from datetime import datetime
import re
from django.contrib.auth.mixins import LoginRequiredMixin
from django.shortcuts import render, redirect
from django.template.loader import render_to_string
from django.urls import reverse_lazy
from django.views import View
from django.views.generic import ListView, DetailView, CreateView, UpdateView, DeleteView, TemplateView
from django.db.models import Count, Q
from django.http import HttpResponseRedirect, JsonResponse, HttpResponse
from django.contrib import messages
from django.contrib.messages.views import SuccessMessageMixin
from django.contrib.auth.mixins import PermissionRequiredMixin

from asplinks import settings
from .forms import AirlineForm, UsAirlineForm, CountryForm
from .models import Airline, Country, State, CommissionHistory, RemoteServers, LatestFiles, FTPhistory

from report.models import ReportPeriod


class HomeView(LoginRequiredMixin, View):
    """Home view for all users."""
    template_name = 'home.html'

    def get(self, request):
        return render(request, self.template_name, {})


class AirlineListView(PermissionRequiredMixin, ListView):
    """Airline listing with pagination."""

    model = Airline
    template_name = 'airline-listing.html'
    context_object_name = 'airlines'

    paginate_by = 20
    permission_required = ('report.view_airline_management',)

    def get_queryset(self):
        order = self.request.GET.get('order_by', 'id')
        qs = Airline.objects.filter(country=self.request.session.get('country')).select_related('product_manager').order_by(order)
        query = self.request.GET.get('q')
        if query:
            qs = qs.filter(
                Q(name__icontains=query) | Q(code__icontains=query) | Q(abrev__icontains=query) | Q(
                    product_manager__first_name__icontains=query)
            )

        return qs

    def get_context_data(self, **kwargs):
        context = super(AirlineListView, self).get_context_data(**kwargs)
        context['activate'] = 'airlines'
        context['order_by'] = self.request.GET.get('order_by', 'id')
        context['query'] = self.request.GET.get('q', '')
        return context


class AirlineDetailsView(PermissionRequiredMixin, DetailView):
    """Airline details view."""

    model = Airline
    template_name = 'airline-details.html'
    context_object_name = 'object'

    permission_required = ('report.view_airline_management',)

    def get_context_data(self, **kwargs):
        context = super(AirlineDetailsView, self).get_context_data(**kwargs)
        context['activate'] = 'airlines'
        return context


class AirlineCreateView(PermissionRequiredMixin, SuccessMessageMixin, CreateView):
    """ Airline creation view."""

    model = Airline
    form_class = AirlineForm
    template_name = 'add-airline.html'
    success_message = "%(name)s was created successfully"
    success_url = '/airlines'
    permission_required = ('report.change_airline_management',)

    def get_success_url(self):
        next = self.request.GET.get('next', None)
        if next:
            return next
        return self.success_url

    def get_form_class(self):
        country = Country.objects.filter(id=self.request.session.get('country')).first()
        if country.name == 'United States':
            return UsAirlineForm
        else:
            return AirlineForm

    def get_form_kwargs(self, *args, **kwargs):
        form_kwargs = super(AirlineCreateView, self).get_form_kwargs(*args, **kwargs)
        form_kwargs['country'] = self.request.session.get('country')
        return form_kwargs


    def get_context_data(self, **kwargs):
        context = super(AirlineCreateView, self).get_context_data(**kwargs)
        context['activate'] = 'airlines'
        context['country'] = Country.objects.filter(id=self.request.session.get('country')).first()
        return context


class AirlineUpdateView(SuccessMessageMixin, PermissionRequiredMixin, UpdateView):
    """Airline edit view."""

    form_class = AirlineForm
    model = Airline
    template_name = 'edit-airline.html'

    context_object_name = 'object'
    success_message = "%(name)s was updated successfully"
    permission_required = ('report.change_airline_management',)

    def get_form_kwargs(self, *args, **kwargs):
        form_kwargs = super(AirlineUpdateView, self).get_form_kwargs(*args, **kwargs)
        form_kwargs['country'] = self.get_object().country.pk
        return form_kwargs

    def get_context_data(self, **kwargs):
        context = super(AirlineUpdateView, self).get_context_data(**kwargs)
        context['activate'] = 'airlines'
        return context


class AirlineDeleteView(PermissionRequiredMixin, DeleteView):
    """Airline deletion view."""

    model = Airline

    success_url = '/airlines'
    permission_required = ('report.change_airline_management',)

    def delete(self, *args, **kwargs):
        self.get_object().delete()
        messages.add_message(self.request, messages.SUCCESS,
                             'Airline deleted successfully.')
        return HttpResponseRedirect(self.success_url)

    def get_success_url(self):
        return reverse_lazy('airlines')


class AirlineCommissionsView(PermissionRequiredMixin, TemplateView):
    """AirlineCommissions listing with pagination and ordering."""

    template_name = 'airline-commissions.html'

    paginate_by = 20
    permission_required = ('report.view_airline_management',)

    def get_context_data(self, **kwargs):
        context = super(AirlineCommissionsView, self).get_context_data(**kwargs)
        context['activate'] = 'airlines'
        # context['add_yqyr'] = 0

        y = CommissionHistory.objects.filter(
            airline=kwargs.get('pk')).order_by('from_date').first()
        # print("--------------------------+++++",y.add_yq_yr)
        s=CommissionHistory.objects.filter(
            airline=kwargs.get('pk'))

        context['add_yqyr'] = False
        if y:
            context['add_yqyr'] = y.add_yq_yr


        context['airline'] = Airline.objects.get(pk=kwargs.get('pk'))
        context['airlineid']=kwargs.get('pk')

        context['commission_histories'] = CommissionHistory.objects.filter(
            airline=kwargs.get('pk')).order_by('from_date')
        context['first_peds'] = ReportPeriod.objects.filter(country=self.request.session['country'], week=1).order_by('ped')
        context['country'] = Country.objects.filter(id=self.request.session.get('country')).first()
        return context

    def post(self, request, *args, **kwargs):

        # if self.request.is_ajax():
        #     pk = kwargs.get('pk')

        #     add_yqyr = self.request.POST.get('yqyr')
        #     if add_yqyr == 'true':
        #         add_yqyr = True
        #     if add_yqyr =='false':
        #         add_yqyr = False


        #     # update= CommissionHistory.objects.filter(
        #     #     airline=pk).update(add_yq_yr=add_yqyr)



        #     return HttpResponseRedirect(reverse_lazy("airline_commissions", kwargs={'pk': pk}))
        # else:
        print("---------------iam not case")
        pk = kwargs.get('pk')
        ped = self.request.POST.get('ped', None)
        to_date = self.request.POST.get('to_date', None)
        commission_rate = self.request.POST.get('commission_rate', None)
        commission_type = self.request.POST.get('commission_type', None)
        airline = Airline.objects.get(pk=pk)
        add_yqyr = self.request.POST.get('add_yqyr')

        if str(add_yqyr) == 'on':
            add_yqyr = True

        if str(add_yqyr) == 'None':
            add_yqyr = False





        if ped:
            from_date = datetime.strptime(ped,'%Y-%m-%d').date()



        if CommissionHistory.objects.filter(airline=airline, type=commission_type, from_date=from_date).exists():
            messages.add_message(self.request, messages.ERROR,
                                'There is already an entry for the same commission type on same from date.')
            return HttpResponseRedirect(reverse_lazy("airline_commissions", kwargs={'pk': pk}))

        CommissionHistory.objects.create(airline=airline,add_yq_yr=add_yqyr, rate=commission_rate, from_date=from_date,type=commission_type, to_date=to_date or None)




        messages.add_message(self.request, messages.SUCCESS,
                            'Commission added successfully.')
        return HttpResponseRedirect(reverse_lazy("airline_commissions", kwargs={'pk': pk}))


class AirlineCommissionDelete(PermissionRequiredMixin, DeleteView):
    """Delete Agency Commission"""

    model = CommissionHistory
    permission_required = ('report.change_airline_management',)

    def delete(self, *args, **kwargs):
        self.get_object().delete()
        messages.add_message(self.request, messages.SUCCESS,
                             'Commission deleted successfully.')
        return HttpResponseRedirect('/airlines')


class ListCountryView(PermissionRequiredMixin, ListView):
    """Country listing with pagination."""

    model = Country
    template_name = 'country-listing.html'
    context_object_name = 'countries'
    permission_required = ('auth.view_group',)

    def get_queryset(self):
        order = self.request.GET.get('order_by', 'name')
        qs = Country.objects.order_by(order)
        query = self.request.GET.get('q')
        if query:
            qs = qs.filter(
                Q(name__icontains=query)
            )
        return qs

    def get_context_data(self, **kwargs):
        context = super(ListCountryView, self).get_context_data(**kwargs)
        context['activate'] = 'users'
        context['order_by'] = self.request.GET.get('order_by', 'id')
        context['query'] = self.request.GET.get('q', '')
        return context


class CountryUpdateView(PermissionRequiredMixin, UpdateView):
    """ Country edit view."""

    form_class = CountryForm
    model = Country
    template_name = 'edit-country.html'
    context_object_name = 'object'
    success_url = '/countries'
    permission_required = ('account.change_airline_management',)


    def get_context_data(self, **kwargs):
        context = super(CountryUpdateView, self).get_context_data(**kwargs)
        return context

    def form_valid(self, form):
        messages.add_message(self.request, messages.SUCCESS,
                             'Country details updated successfully.')
        return super(CountryUpdateView, self).form_valid(form)



class AddCountryView(CreateView):
    """ Country creation view."""

    model = Country
    template_name = 'add-country.html'
    success_message = "%(name)s was created successfully"
    permission_required = ('account.change_airline_management',)
    form_class = CountryForm
    success_url = '/countries'

    def form_valid(self, form):
        self.object = form.save()
        self.object.save()
        messages.add_message(self.request, messages.SUCCESS,
                             'Country added successfully')
        return HttpResponseRedirect(self.get_success_url())

    def get_context_data(self, **kwargs):
        context = super(AddCountryView, self).get_context_data(**kwargs)
        return context


class SetCountryView(LoginRequiredMixin, View):
    """

    Country switch view
    """

    def post(self, request, *args, **kwargs):
        try:
            country = request.POST.get('country')
            self.request.session['country'] = country
            return JsonResponse({"message": "switched country"})
        except Exception as e:
            return JsonResponse(
                    {"message": "couldn't switch Country"}, status=400)


class FTPManagementView(TemplateView):
    """
        FTPManagementView allows to connect server, go through the directories and download files.
    """
    template_name = 'ftp-management.html'

    def get(self,request):

        ajax_type = request.GET.get("type")
        if request.is_ajax():
            response = {}
            if ajax_type == "connect-server":
                host = request.GET.get("host")
                user = request.GET.get("user")
                password = request.GET.get("password")
                selected_r_host = request.GET.get("selected_r_host")
                if request.GET.get("from-inp") == "False":
                    server_obj = RemoteServers.objects.get(id=selected_r_host)
                    host = server_obj.hostname
                    user = server_obj.user
                    password = server_obj.password
                current_path = "/"+user
                prev_path = None
                file_list = connect_server(host,user,password,current_path)
                response['html_content'] = render_to_string("includes/ftp-management-table-content.html",{
                    'file_list':file_list,
                    'current_path':current_path,
                    'server_name':user+"@"+host,
                    # 'r_server_id':server_obj.id,
                    "prev_path":prev_path
                })
            elif ajax_type == "forward-directory":
                host = request.GET.get("host")
                user = request.GET.get("user")
                password = request.GET.get("password")
                selected_file = request.GET.get("selected_file")
                current_path = request.GET.get("current_path")
                new_current_path = current_path+selected_file
                file_list = connect_server(host, user, password,new_current_path)
                prev_path = current_path
                response['html_content'] = render_to_string("includes/ftp-management-table-content.html", {
                    'file_list': file_list,
                    'current_path': new_current_path,
                    'server_name': user + "@" + host,
                    "prev_path": prev_path
                })

            elif ajax_type == "backward-directory":
                host = request.GET.get("host")
                user = request.GET.get("user")
                password = request.GET.get("password")
                current_path = request.GET.get("prev_path")
                if current_path == "/"+user:
                    prev_path = None
                else:
                    prev_path = str(current_path).split("/")[0]
                file_list = connect_server(host, user, password, current_path)
                response['html_content'] = render_to_string("includes/ftp-management-table-content.html", {
                    'file_list': file_list,
                    'current_path': current_path,
                    'server_name': user + "@" + host,
                    "prev_path": prev_path
                })

            elif ajax_type == "download-file":
                host = request.GET.get("host")
                user = request.GET.get("user")
                password = request.GET.get("password")
                selected_file = request.GET.get("selected_file")
                current_path = request.GET.get("current_path")
                current_path = current_path+"/"+selected_file
                target_path = "{}Downloads/{}".format(settings.MEDIA_ROOT, selected_file)
                resp_target_path  = "/media/Downloads/{}".format(selected_file)
                file_list = connect_server(host, user, password, current_path,filename=selected_file,is_download=True,target_path=target_path)
                response['target_path'] = resp_target_path
            return HttpResponse(json.dumps(response), content_type="application/json")

        else:
            context = {}
            historylist = FTPhistory.objects.filter(ftp_obj__countrycode=request.session.get("country")).order_by("-created_on")
            context['historylist'] = historylist
            return render(request,self.template_name,context)

class AddRemoteHostView(TemplateView):
    """
        AddRemoteHostView allows to manage hosts.
    """
    template_name = 'add-remote-host.html'
    def get(self,request):

        if request.is_ajax():
            response = {}
            ajax_type = request.GET.get("type")
            if ajax_type == "add-remote-server":
                host = request.GET.get("host")
                user = request.GET.get("user")
                password = request.GET.get("password")
                port = request.GET.get("port")
                countrycode = request.session.get("country")
                if RemoteServers.objects.filter(hostname=host,user=user,password=password,port=port,countrycode=countrycode):
                    response['status'] = False
                else:
                    RemoteServers.objects.create(hostname=host, user=user, password=password,port=port,countrycode=countrycode)
                    response['status'] = True
            elif ajax_type == "delete-remote-server":
                RemoteServers.objects.get(id=request.GET.get("server_id")).delete()
                response['status'] = True
            return HttpResponse(json.dumps(response), content_type="application/json")
        else:
            context = {}
            context['remote_hosts'] = RemoteServers.objects.filter(countrycode=request.session.get("country"))
        return render(request, self.template_name,context)


def country_import():
    countries = [
        {'name': 'Canada', 'code': 'CA', 'states': [
            ('Ontario', 'ON'),
            ('Quebec', 'QC'),
            ('Nova Scotia', 'NS'),
            ('New Brunswick', 'NB'),
            ('Manitoba', 'MB'),
            ('British Columbia', 'BC'),
            ('Prince Edward Island', 'PE'),
            ('Saskatchewan', 'SK'),
            ('Alberta', 'AB'),
            ('Newfoundland and Labrador', 'NL'),
            ('Yukon', 'YT'),
            ('Northwest Territories', 'NT'),
            ('Nunavut', 'NU'),
            ('Canada', 'CA'),
        ]},
        {'name': 'United States', 'code': 'US', 'states': [
            ('Alabama', 'AL'),
            ('Alaska', 'AK'),
            ('Arizona', 'AZ'),
            ('Arkansas', 'AR'),
            ('California', 'CA'),
            ('Colorado', 'CO'),
            ('Connecticut', 'CT'),
            ('Delaware', 'DE'),
            ('Florida', 'FL'),
            ('Georgia', 'GA'),
            ('Hawaii', 'HI'),
            ('Idaho', 'ID'),
            ('Illinois', 'IL'),
            ('Indiana', 'IN'),
            ('Iowa', 'IA'),
            ('Kansas', 'KS'),
            ('Kentucky', 'KY'),
            ('Louisiana', 'LA'),
            ('Maine', 'ME'),
            ('Maryland ', 'MD'),
            ('Massachusett', 'MA'),
            ('Michigan', 'MI'),
            ('Minnesota', 'MN'),
            ('Mississippi', 'MS'),
            ('Missouri', 'MO'),
            ('Montana', 'MT'),
            ('Nebraska', 'NE'),
            ('Nevada', 'NV'),
            ('New Hampshire', 'NH'),
            ('New Jersey', 'NJ'),
            ('New Mexico', 'NM'),
            ('New York', 'NY'),
            ('North Carolina', 'NC'),
            ('North Dakota', 'ND'),
            ('Ohio', 'OH'),
            ('Oklahoma', 'OK'),
            ('Oregon', 'OR'),
            ('Pennsylvania', 'PA'),
            ('Rhode Island', 'RI'),
            ('South Carolina', 'SC'),
            ('South Dakota', 'SD'),
            ('Tennessee', 'TN'),
            ('Texas', 'TX'),
            ('Utah', 'UT'),
            ('Vermont', 'VT'),
            ('Virginia', 'VA'),
            ('Washington', 'WA'),
            ('West Virginia', 'WV'),
            ('Wisconsin', 'WI'),
            ('Wyoming', 'WY'),
            ('Puerto Rico', 'PR'),
            ('Virgin Islands', 'VI'),
            ('Guam', 'GU'),
            ('Northern Mariana Islands', 'MP'),
            ('American Samoa', 'AS'),
            ('District of Columbia', 'DC'),
            ('Federated States of Micronesia', 'FM'),
            ('Marshall Islands', 'MH'),
            ('Palau', 'PW'),
            ('Armed Forces Africa', 'AE'),
            ('Armed Forces Americas', 'AA'),
            ('Armed Forces Canada', 'AE'),
            ('Armed Forces Europe', 'AE'),
            ('Armed Forces Middle East', 'AE'),
            ('Armed Forces Pacific', 'AP'),
        ]},
    ]

    for country in countries:
        c, created = Country.objects.get_or_create(
                name=country.get('name'),
                defaults={
                    'code': country.get('code'),
                }
            )
        for state in country.get('states'):
            st, created = State.objects.get_or_create(
                name=state[0],
                defaults={
                    'abrev': state[1],
                    'country': c
                }
            )


def connect_server(host,user,password,current_path=None,is_download=False,filename=None,target_path=None):
    from paramiko import SSHClient
    from paramiko import AutoAddPolicy
    from datetime import datetime
    ssh = SSHClient()
    ssh.load_system_host_keys()
    ssh.set_missing_host_key_policy(AutoAddPolicy())
    ssh.connect(host, username=user, password=password)

    ftp = ssh.open_sftp()
    file_list = []

    latest = 0
    latestfile = None

    for fileattr in ftp.listdir_attr():
        if fileattr.st_mtime > latest:
            latest = fileattr.st_mtime
            latestfile = fileattr.filename


    if is_download:
        ftp.get(current_path,target_path)
    else:
        if current_path:
            ftp.chdir(current_path)

        files = ftp.listdir_attr()
        for file in files:
            info = file.st_size
            timestamp = ftp.stat(file.filename).st_mtime
            dt_object = datetime.fromtimestamp(timestamp)
            date_time_str = dt_object.strftime("%d/%m/%y %-H:%M:%S")
            import stat

            file_dict = {
                'filename': file.filename,
                'modified_date': date_time_str,
                'size': str(info) + " B",
                "color": "green" if stat.S_ISDIR(file.st_mode) else "#4D4E5E",
                "is_file": "True" if stat.S_ISDIR(file.st_mode) else "False"
            }
            file_list.append(file_dict)

    ftp.close()
    ssh.close()
    return file_list


def download_latest(host,user,password,countrycode,target_path=None,initial=True,latest=0,ftp_obj=None):
    from paramiko import SSHClient
    from paramiko import AutoAddPolicy
    import logging
    ssh = SSHClient()
    ssh.load_system_host_keys()
    ssh.set_missing_host_key_policy(AutoAddPolicy())
    # key_filename = '{}testserver.pem'.format(settings.MEDIA_ROOT)
    ssh.connect(host, username=user, password=password,allow_agent=False)
    import stat


    latesttime = latest
    current_path = "/" + user+"/"
    # current_path = "/home/ubuntu/report_files/uploadedfiles/"
    import requests
    client = requests.session()
    fetch_url = settings.BASE_DOMAIN
    url = "{}reports/scheduler-report-upload/".format(settings.BASE_DOMAIN)
    client.get(fetch_url)
    print("client  ",client)
    csrftoken = client.cookies['csrftoken']
    print("csrftoken  ",csrftoken)
    headers = {'X-CSRFToken': csrftoken}
    cookies = {'csrftoken': csrftoken}
    ssh.exec_command(current_path)
    stdin, stdout, stderr = ssh.exec_command(current_path)
    # print("_----___--___---__---__--__-         ",stdin,stdout,stderr)
    ftp = ssh.open_sftp()
    print("C DIR     ",ftp.getcwd())
    ftp.chdir(current_path)
    print("---->>>    C DIR     ", ftp.getcwd())
    logging.info("---->>>    C DIR    "+ftp.getcwd())
    for fileattr in ftp.listdir_attr():
        print(latesttime)
        # print("fileattr   ",fileattr,stat.S_ISDIR(fileattr.st_mode))
        if not stat.S_ISDIR(fileattr.st_mode):
            filename = fileattr.filename
            print("filename        ",filename)
            try:
                FTPhistory.objects.get(ftp_obj=ftp_obj,file=filename,status=True)
                print(filename,"     --->>>  ALREADY EXISTS")
            except:
                # try:
                if not str(filename).lower().endswith(('.csv')):
                    file_path = current_path+filename
                    print(filename,"     --->>>  NEW FILE")
                    # print("file_path  ",initial,fileattr.st_mtime,"   ---   ",latesttime,file_path,ftp.getcwd())
                    # logging.info("file_path  "+"    "+str(fileattr.st_mtime)+ "  ---   "+str(latesttime)+"    "+str(file_path)+"    "+str(ftp.getcwd()))
                    target_path = "{}Downloads/{}".format(settings.MEDIA_ROOT, filename)
                    if fileattr.st_mtime > int(latesttime):
                        latesttime = fileattr.st_mtime
                    if not initial:
                        print("=====>>  >>  >>",initial,fileattr.st_mtime)
                        # logging.info("NOT INITIAL")
                        if fileattr.st_mtime > int(latest) or True:
                            ftp.get(file_path, target_path)
                            files = {'file': open(target_path, 'rb')}
                            payload = {

                                'from_scheduler':True,
                                'countrycode':countrycode,
                                'csrfmiddlewaretoken': csrftoken
                            }
                            resp = requests.post(url, data=payload, files=files, headers=headers, cookies=cookies)
                            print(">>>            ",resp)
                            logging.info("RESPONSEEEEEE"+resp)
                    else:
                        # print("_________",initial,fileattr.st_mtime > int(latest))
                        print("FILE PATH    --->>>   ",file_path)
                        ftp.get(file_path, target_path)
                        payload = {

                            'from_scheduler': True,
                            'countrycode': countrycode,
                            'csrfmiddlewaretoken': csrftoken
                        }
                        files = {'file': open(target_path, 'rb')}
                        resp = requests.post(url, data=payload,files=files,headers=headers,cookies=cookies)
                        print("RESPPP           ",resp)
                    FTPhistory.objects.get_or_create(ftp_obj=ftp_obj,file=filename,status=True)
                # except Exception as e:
                #     print ("Error Occurred :",e)
                #     FTPhistory.objects.get_or_create(ftp_obj=ftp_obj, file=filename, status=False)
                    logging.info("======================================================================================")
    if initial:
        lf_obj = LatestFiles.objects.create(latest=latesttime, ftp_obj=ftp_obj)
    else:
        try:
            lf_obj = LatestFiles.objects.get(ftp_obj=ftp_obj)
            lf_obj.latest = latesttime
            lf_obj.save()
        except:
            lf_obj = LatestFiles.objects.create(latest=latesttime, ftp_obj=ftp_obj)
