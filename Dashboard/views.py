from django.shortcuts import render
from django.http import HttpResponse,  JsonResponse
from django.core.serializers.json import DjangoJSONEncoder
from django.views import View
from django.views.generic import FormView,View,ListView
from django.views import generic
from Dashboard.forms import Dashboardform
from Dashboard.models import DashboardModel
from django.conf import settings
import datetime
import time
import sys
from django.core.files.storage import FileSystemStorage
import os
from django.urls import reverse_lazy
import pandas as pd
from pandas import MultiIndex
from django.template.loader import render_to_string
import json
import datetime
from django.core import serializers
from main.models import Country
import numpy as np
from django.db.models import Sum
import openpyxl
import plotly.offline as py
import plotly.graph_objs as go
from django.contrib.auth.models import Group
from django.contrib.auth.mixins import PermissionRequiredMixin

class Dashboard(View):
    def get(self, request):
        # DashboardModel.objects.filter(Country="COTE DE IVOIRE").delete()
        # DashboardModel.objects.all().delete()
        airline = request.GET.getlist('airline_dropdown')
        agreement_type = request.GET.getlist('agreement_dropdown')
        year = request.GET.get("year")
        month = request.GET.get('month')
        country = request.GET.get('country_dropdown')
        country__ = 'United States' if country == "USA" else country
        cntry_object = Country.objects.filter(name__iexact = country__)
        selected_flag = None
        if cntry_object:
            selected_flag = cntry_object.first().flag
        user_admin_flag = False
        currnt_usr = request.user
        grps = currnt_usr.groups.all()
        grp_name = [grp.name for grp in grps]
        global_user = ["aasokan@airlinepros.net","amathew@airlinepros.com","vbaby@airlinepros.net","mjanardanan@airlinepros.net",
                       "sudhak@airlinepros.net","dpurushothaman@airlinepros.net","zaboobacker@airlinepros.net","snanavati@airlinepros.com",
                       "lbarber@airlinepros.net","admin@abhilash.com"]
        if currnt_usr.email in global_user:
            user_admin_flag = True
        if currnt_usr.email in global_user:
            pd.set_option('display.float_format', '{:.2f}'.format)
            df = pd.DataFrame(list(DashboardModel.objects.all().values()))
            month_order = ['JANUARY', 'FEBRUARY', 'MARCH', 'APRIL', 'MAY', 'JUNE', 'JULY', 'AUGUST', 'SEPTEMBER', 'OCTOBER', 'NOVEMBER', 'DECEMBER']
            month_dtype = pd.api.types.CategoricalDtype(categories=month_order, ordered=True)
            df['Month'] = df['Month'].astype(month_dtype)
            global_qs = DashboardModel.objects.all()
            
            try:
                if country:
                  
                    global_qs = global_qs.filter(Country=country)
                    df = pd.DataFrame(list(global_qs.values()))
                    month_order = ['JANUARY', 'FEBRUARY', 'MARCH', 'APRIL', 'MAY', 'JUNE', 'JULY', 'AUGUST', 'SEPTEMBER', 'OCTOBER', 'NOVEMBER', 'DECEMBER']
                    month_dtype = pd.api.types.CategoricalDtype(categories=month_order, ordered=True)
                    df['Month'] = df['Month'].astype(month_dtype)
                    
            except Exception as e:
                print(e,".....................")
                            
            # if year:
            #     print("year filterrrrrrrrrrrrrrrrrrrrrrr")
            #     global_qs = global_qs.filter(Year=year)
            #     df = pd.DataFrame(list(global_qs.values()))
            #     print(df,"//////////////////")
                    
                
            
            try:
                if agreement_type:
                  
                    global_qs = global_qs.filter(Agreement_Type__in=agreement_type)
                    df = pd.DataFrame(list(global_qs.values()))
                    month_order = ['JANUARY', 'FEBRUARY', 'MARCH', 'APRIL', 'MAY', 'JUNE', 'JULY', 'AUGUST', 'SEPTEMBER', 'OCTOBER', 'NOVEMBER', 'DECEMBER']
                    month_dtype = pd.api.types.CategoricalDtype(categories=month_order, ordered=True)
                    df['Month'] = df['Month'].astype(month_dtype)
                   
                  
    
            except Exception as e:
                print(e,".....................")
            try:   
                if airline:
                    global_qs = global_qs.filter(Airline__in=airline)
                    df = pd.DataFrame(list(global_qs.values()))
                    month_order = ['JANUARY', 'FEBRUARY', 'MARCH', 'APRIL', 'MAY', 'JUNE', 'JULY', 'AUGUST', 'SEPTEMBER', 'OCTOBER', 'NOVEMBER', 'DECEMBER']
                    month_dtype = pd.api.types.CategoricalDtype(categories=month_order, ordered=True)
                    df['Month'] = df['Month'].astype(month_dtype)
                   
            except Exception as e:
                print(e,".....................")
                
            try:
              
                if month:
                    df = pd.DataFrame(list(DashboardModel.objects.filter(Month=month).values()))
                    month_order = ['JANUARY', 'FEBRUARY', 'MARCH', 'APRIL', 'MAY', 'JUNE', 'JULY', 'AUGUST', 'SEPTEMBER', 'OCTOBER', 'NOVEMBER', 'DECEMBER']
                    month_dtype = pd.api.types.CategoricalDtype(categories=month_order, ordered=True)
                    df['Month'] = df['Month'].astype(month_dtype)
            except Exception as e:
                print(e,".....................")
    
            # by month and year
            
            try:
                month_order = ['JANUARY', 'FEBRUARY', 'MARCH', 'APRIL', 'MAY', 'JUNE', 'JULY', 'AUGUST', 'SEPTEMBER', 'OCTOBER', 'NOVEMBER', 'DECEMBER']
                month_dtype = pd.api.types.CategoricalDtype(categories=month_order, ordered=True)
                df['Month'] = df['Month'].astype(month_dtype)
                gross = df.pivot_table( index="Month",columns=["Year"],values=["Gross"],aggfunc="sum")
                nett = df.pivot_table( index="Month",columns=["Year"],values=["Nett"],aggfunc="sum")
                label = list(gross.index)
                
                gross_year = MultiIndex.from_tuples(gross.columns)
                values = gross.values
                Gross_values = {}
                for i, year in enumerate(gross_year.get_level_values(1)):
                    Gross_values[year] = [row[i] for row in values]
                    
                nett_year = MultiIndex.from_tuples(nett.columns)
                values = nett.values
                Nett_values = {}
                for i, year in enumerate(nett_year.get_level_values(1)):
                    Nett_values[year] = [row[i] for row in values]

                # by airline
                gross = (df.groupby(["Airline"], sort=False)["Gross"].agg("sum"))
                label_airline = list(gross.index)
                values_gross_airline = list(gross.values)
                nett = df.groupby(["Airline"],sort=False)["Nett"].agg("sum")
                values_nett_airline = list(nett.values)

                # by year
                gross = (df.groupby(["Year"], sort=False)["Gross"].agg("sum"))
                label_year = list(gross.index)
                values_gross_year = list(gross.values)
                nett = df.groupby(["Year"],sort=False)["Nett"].agg("sum")
                values_nett_year = list(nett.values)
            
            
            
                unique_year = DashboardModel.objects.all().values_list('Year',flat=True).distinct().order_by()
                if country:
                    unique_agreement_type = DashboardModel.objects.filter(Country =country ).values_list('Agreement_Type',flat=True).distinct().order_by('Agreement_Type')
                else:
                    unique_agreement_type = DashboardModel.objects.all().values_list('Agreement_Type',flat=True).distinct().order_by()
                if country:
                    unique_airline = DashboardModel.objects.filter(Country =country ).values_list('Airline',flat=True).distinct().order_by('Airline')
                else:
                    unique_airline = DashboardModel.objects.all().values_list('Airline',flat=True).distinct().order_by()
                unique_month = DashboardModel.objects.all().values_list('Month',flat=True).distinct().order_by()
                unique_country = DashboardModel.objects.all().values_list('Country',flat=True).distinct().order_by('Country')
                return render(request, 'dashboard.html',{'no_data':False,'label':label,'Gross_values':Gross_values,'Nett_values':Nett_values,
                                                        'label_airline':label_airline,'values_gross_airline':values_gross_airline,'values_nett_airline':values_nett_airline,
                                                        'label_year':label_year,'values_gross_year':values_gross_year,'values_nett_year':values_nett_year,
                                                        'unique_year':unique_year,'unique_agreement_type':unique_agreement_type,'unique_airline':unique_airline,
                                                        'unique_month':unique_month,'country_dropdown':country ,'airline_dropdown': airline, 'agreement_dropdown' : agreement_type,'unique_country':unique_country,'user_admin_flag':user_admin_flag,'selected_flag':selected_flag,
                                                        })
            except Exception as e:
                
                print(e,".....................")
                unique_year = DashboardModel.objects.all().values_list('Year',flat=True).distinct().order_by()
                if country:
                    unique_agreement_type = DashboardModel.objects.filter(Country =country ).values_list('Agreement_Type',flat=True).distinct().order_by('Agreement_Type')
                else:
                    unique_agreement_type = DashboardModel.objects.all().values_list('Agreement_Type',flat=True).distinct().order_by()
                if country:
                    unique_airline = DashboardModel.objects.filter(Country =country ).values_list('Airline',flat=True).distinct().order_by('Airline')
                else:
                    unique_airline = DashboardModel.objects.all().values_list('Airline',flat=True).distinct().order_by()
                unique_month = DashboardModel.objects.all().values_list('Month',flat=True).distinct().order_by()
                unique_country = DashboardModel.objects.all().values_list('Country',flat=True).distinct().order_by('Country')
                return render(request, 'dashboard.html',{'no_data':True,'label':None,'Gross_values':None,'Nett_values':None,
                                                        'label_airline':None,'values_gross_airline':None,'values_nett_airline':None,
                                                        'label_year':None,'values_gross_year':None,'values_nett_year':None,
                                                        'unique_year':unique_year,'unique_agreement_type':unique_agreement_type,'unique_airline':unique_airline,
                                                        'unique_month':unique_month, 'country_dropdown':country,'airline_dropdown': airline, 'agreement_dropdown' : agreement_type,'unique_country':unique_country,'user_admin_flag':user_admin_flag
                                                        })
                

        else:
            choosen_country_ = request.session.get('country')
            obj = Country.objects.get(id = choosen_country_)
           
        
            if obj.name == "United States":
                country_var= "USA"
            else:
                country_var = obj.name
            



            pd.set_option('display.float_format', '{:.2f}'.format)
            df = pd.DataFrame(list(DashboardModel.objects.filter(Country__iexact=country_var).values()))
            month_order = ['JANUARY', 'FEBRUARY', 'MARCH', 'APRIL', 'MAY', 'JUNE', 'JULY', 'AUGUST', 'SEPTEMBER', 'OCTOBER', 'NOVEMBER', 'DECEMBER']
            month_dtype = pd.api.types.CategoricalDtype(categories=month_order, ordered=True)
            df['Month'] = df['Month'].astype(month_dtype)

            global_qs = DashboardModel.objects.all()  
            try:  
                if country:
                    global_qs = global_qs.filter(Country=country)
                    df = pd.DataFrame(list(DashboardModel.objects.filter( Country__iexact=country_var, Country=country).values()))
                    month_order = ['JANUARY', 'FEBRUARY', 'MARCH', 'APRIL', 'MAY', 'JUNE', 'JULY', 'AUGUST', 'SEPTEMBER', 'OCTOBER', 'NOVEMBER', 'DECEMBER']
                    month_dtype = pd.api.types.CategoricalDtype(categories=month_order, ordered=True)
                    df['Month'] = df['Month'].astype(month_dtype)
            except Exception as e:
                print(e,"...............")
            try:                  
                if year:
                 
                    global_qs = global_qs.filter(Country__iexact=country_var, Year=year)
                    df = pd.DataFrame(list(global_qs.values()))
                    month_order = ['JANUARY', 'FEBRUARY', 'MARCH', 'APRIL', 'MAY', 'JUNE', 'JULY', 'AUGUST', 'SEPTEMBER', 'OCTOBER', 'NOVEMBER', 'DECEMBER']
                    month_dtype = pd.api.types.CategoricalDtype(categories=month_order, ordered=True)
                    df['Month'] = df['Month'].astype(month_dtype)
              
            except Exception as e:
                print(e,".................")

            
            try:
                if agreement_type:
                 
                    global_qs = global_qs.filter(Country__iexact=country_var ,Agreement_Type__in=agreement_type)
                    df = pd.DataFrame(list(global_qs.values()))
                    month_order = ['JANUARY', 'FEBRUARY', 'MARCH', 'APRIL', 'MAY', 'JUNE', 'JULY', 'AUGUST', 'SEPTEMBER', 'OCTOBER', 'NOVEMBER', 'DECEMBER']
                    month_dtype = pd.api.types.CategoricalDtype(categories=month_order, ordered=True)
                    df['Month'] = df['Month'].astype(month_dtype)
                   

            except Exception as e:
                print(e,".................")

            try:
                if airline:
                    
                    global_qs = global_qs.filter(Country__iexact=country_var ,Airline__in=airline)
                    df = pd.DataFrame(list(global_qs.values()))
                    month_order = ['JANUARY', 'FEBRUARY', 'MARCH', 'APRIL', 'MAY', 'JUNE', 'JULY', 'AUGUST', 'SEPTEMBER', 'OCTOBER', 'NOVEMBER', 'DECEMBER']
                    month_dtype = pd.api.types.CategoricalDtype(categories=month_order, ordered=True)
                    df['Month'] = df['Month'].astype(month_dtype)
            except Exception as e:
                print(e,"........................")

            try:
                if month:
                   
                    df = pd.DataFrame(list(DashboardModel.objects.filter(Country__iexact=country_var ,Month=month).values()))
                    month_order = ['JANUARY', 'FEBRUARY', 'MARCH', 'APRIL', 'MAY', 'JUNE', 'JULY', 'AUGUST', 'SEPTEMBER', 'OCTOBER', 'NOVEMBER', 'DECEMBER']
                    month_dtype = pd.api.types.CategoricalDtype(categories=month_order, ordered=True)
                    df['Month'] = df['Month'].astype(month_dtype)
            except Exception as e:
                print(e,".............................")  

            # by month and year
            
            try:
                month_order = ['JANUARY', 'FEBRUARY', 'MARCH', 'APRIL', 'MAY', 'JUNE', 'JULY', 'AUGUST', 'SEPTEMBER', 'OCTOBER', 'NOVEMBER', 'DECEMBER']
                month_dtype = pd.api.types.CategoricalDtype(categories=month_order, ordered=True)
                df['Month'] = df['Month'].astype(month_dtype)
                gross = df.pivot_table( index="Month",columns=["Year"],values=["Gross"],aggfunc="sum")
                nett = df.pivot_table( index="Month",columns=["Year"],values=["Nett"],aggfunc="sum")
                label = list(gross.index)

                gross_year = MultiIndex.from_tuples(gross.columns)
                values = gross.values
                Gross_values = {}
                for i, year in enumerate(gross_year.get_level_values(1)):
                    Gross_values[year] = [row[i] for row in values]

                nett_year = MultiIndex.from_tuples(nett.columns)
                values = nett.values
                Nett_values = {}
                for i, year in enumerate(nett_year.get_level_values(1)):
                    Nett_values[year] = [row[i] for row in values]

                # by airline
                gross = (df.groupby(["Airline"], sort=False)["Gross"].agg("sum"))
                label_airline = list(gross.index)
                values_gross_airline = list(gross.values)
                nett = df.groupby(["Airline"],sort=False)["Nett"].agg("sum")
                values_nett_airline = list(nett.values)

                # by year
                gross = (df.groupby(["Year"], sort=False)["Gross"].agg("sum"))
                label_year = list(gross.index)
                values_gross_year = list(gross.values)
                nett = df.groupby(["Year"],sort=False)["Nett"].agg("sum")
                values_nett_year = list(nett.values)



                unique_year = DashboardModel.objects.filter(Country__iexact=country_var).values_list('Year',flat=True).distinct().order_by()
                unique_agreement_type = DashboardModel.objects.filter(Country__iexact=country_var).values_list('Agreement_Type',flat=True).distinct()
                unique_airline = DashboardModel.objects.filter(Country__iexact=country_var).values_list('Airline',flat=True).distinct().order_by()
                unique_month = DashboardModel.objects.filter(Country__iexact=country_var).values_list('Month',flat=True).distinct().order_by()
                unique_country = DashboardModel.objects.filter(Country__iexact=country_var).values_list('Country',flat=True).distinct().order_by()
                return render(request, 'dashboard.html',{'no_data':False,'label':label,'Gross_values':Gross_values,'Nett_values':Nett_values,
                                                        'label_airline':label_airline,'values_gross_airline':values_gross_airline,'values_nett_airline':values_nett_airline,
                                                        'label_year':label_year,'values_gross_year':values_gross_year,'values_nett_year':values_nett_year,
                                                        'unique_year':unique_year,'unique_agreement_type':unique_agreement_type,'unique_airline':unique_airline,
                                                        'unique_month':unique_month, 'airline_dropdown': airline, 'country_dropdown':country,'agreement_dropdown' : agreement_type,'unique_country':unique_country,'user_admin_flag':user_admin_flag
                                                        })
            
            except Exception as e:
                
                print(e,".....................")
                unique_year = DashboardModel.objects.filter(Country__iexact=country_var).values_list('Year',flat=True).distinct().order_by()
                unique_agreement_type = DashboardModel.objects.filter(Country__iexact=country_var).values_list('Agreement_Type',flat=True).distinct()
                unique_airline = DashboardModel.objects.filter(Country__iexact=country_var).values_list('Airline',flat=True).distinct().order_by()
                unique_month = DashboardModel.objects.filter(Country__iexact=country_var).values_list('Month',flat=True).distinct().order_by()
                unique_country = DashboardModel.objects.filter(Country__iexact=country_var).values_list('Country',flat=True).distinct().order_by()
                return render(request, 'dashboard.html',{'no_data':True,  'label':None,'Gross_values':None,'Nett_values':None,
                                                        'label_airline':None,'values_gross_airline':None,'values_nett_airline':None,
                                                        'label_year':None,'values_gross_year':None,'values_nett_year':None,
                                                        'unique_year':unique_year,'unique_agreement_type':unique_agreement_type,'unique_airline':unique_airline,
                                                        'unique_month':unique_month, 'country_dropdown':country,'airline_dropdown': airline, 'agreement_dropdown' : agreement_type,'unique_country':unique_country,'user_admin_flag':user_admin_flag
                                                        })
       

import threading      
class Uploaddashboardfile(FormView):

    template_name = 'dashboard-file-upload.html'
    form_class = Dashboardform
    # success_url = 'Airlinedashboard/'
    success_url = reverse_lazy('dashboard')
    def form_valid(self, form):
        
        file = self.request.FILES['file']
        file_r = self.request.FILES['file'].read()
        file_name, extention = file.name.split('.')
        media_root = getattr(settings, 'MEDIA_ROOT')
        if extention.lower() == 'xlsx':
            errfiles = []
            errMsg = ''
            zipfileCnt = 0

            media_root = getattr(settings, 'MEDIA_ROOT')
            fs = FileSystemStorage(location=os.path.join(media_root, "Dashboard files/"))
            filename = fs.save(file_name + '.' + extention, file)
            w_filename = file_name+"."+extention
            
            def database_upload_thread(self, file_, w_filename):
                # dataframe = pd.read_excel(file,"DATABASE")
                media_root = getattr(settings, 'MEDIA_ROOT')
                
                
                file_ = media_root+"Dashboard files/"+w_filename
               
                dataframe = pd.read_excel(file_)
              
                
                dataframe = dataframe.fillna(0.0)
                for i in range(len(dataframe)):
                    region = dataframe.loc[i].REGION
                    country = dataframe.loc[i].COUNTRY
                    year = dataframe.loc[i].YEAR
                    airline = dataframe.loc[i].AIRLINE
                    airline_iata_code = dataframe.loc[i]["AIRLINE IATA CODE"]               
                    iata_numeric_code = dataframe.loc[i]["IATA NUMERIC CODE"]
                    agreement_type = dataframe.loc[i]["AGREEMENT TYPE"]
                    contracted_with = dataframe.loc[i]["CONTRACED WITH"]
                    month = dataframe.loc[i].MONTH
                    orc_rate = dataframe.loc[i].ORC_RATE        
                    gross = dataframe.loc[i].GROSS
                    nett = dataframe.loc[i].NETT
                    if "ORC ACTUAL" in dataframe.columns:
                        orc_actual = dataframe.loc[i]["ORC ACTUAL"]
                    else:
                        orc_actual = dataframe.loc[i]["ORC"]
                
                    
                    create =  DashboardModel.objects.get_or_create(Region=region,Country=country,Year=year,Airline=airline,Airline_IATA_Code=airline_iata_code,
                                                                IATA_Numeric_Code=iata_numeric_code,Agreement_Type=agreement_type,Contraced_With=contracted_with,
                                                                Month=month,ORC_Rate=orc_rate,Gross=gross,Nett=nett,ORC_Actual=orc_actual
                                                                )   
            process = threading.Thread(target = database_upload_thread, args = (self, file, w_filename))
            process.start()
          
            text_file = os.path.join(media_root, "Dashboard files/" + filename)
        
        return super(Uploaddashboardfile, self).form_valid(form)
       
    def form_invalid(self, form):
        print("<===============FORM INVALID========================>",form.errors)
        
        
def filter(request, airline_IATA_code =None, region =None, year =None, month =None, country =None, airline =None, IATA_numeric_code =None, agreement_type =None, revenue =None):
    
    currnt_usr = request.user
    grps = currnt_usr.groups.all()
    grp_name = [grp.name for grp in grps]

    
    global_user = ["aasokan@airlinepros.net","amathew@airlinepros.com","vbaby@airlinepros.net","mjanardanan@airlinepros.net",
                       "sudhak@airlinepros.net","dpurushothaman@airlinepros.net","zaboobacker@airlinepros.net","snanavati@airlinepros.com",
                       "lbarber@airlinepros.net","admin@abhilash.com"]

    if currnt_usr.email in global_user:
       
        reset = False
        source_data = DashboardModel.objects.all()        
            
        
    else:   
        
        choosen_country_ = request.session.get('country')
        obj = Country.objects.get(id = choosen_country_)
       
            
        reset = False
        if obj.name == "USA" or obj.name == "United States":
           
            source_data = DashboardModel.objects.all().filter(Country__iexact= 'USA')
        else:
           
            source_data = DashboardModel.objects.all().filter(Country__iexact= obj.name)    
    
    try:       
        if airline_IATA_code:
           
            source_data = source_data.filter(Airline_IATA_Code__in=airline_IATA_code)
            df = pd.DataFrame(source_data.values())
            month_order = ['JANUARY', 'FEBRUARY', 'MARCH', 'APRIL', 'MAY', 'JUNE', 'JULY', 'AUGUST', 'SEPTEMBER', 'OCTOBER', 'NOVEMBER', 'DECEMBER']
            month_dtype = pd.api.types.CategoricalDtype(categories=month_order, ordered=True)
            df['Month'] = df['Month'].astype(month_dtype)
           
            pivot_data = df.pivot_table( index="Month",columns="Year",values=["Gross","Nett"],aggfunc="sum")
            pivot_data = pivot_data.applymap(lambda x: "{:,.2f}".format(x))
            pivot_data = pivot_data.drop(columns=[('Nett',          'Total'),('Gross',          'Total')])

    except Exception as e:
        print(e,"......................")
        
    try:    
        if year:
           
            source_data = source_data.filter(Year__in = year)
            df = pd.DataFrame(source_data.values())
            month_order = ['JANUARY', 'FEBRUARY', 'MARCH', 'APRIL', 'MAY', 'JUNE', 'JULY', 'AUGUST', 'SEPTEMBER', 'OCTOBER', 'NOVEMBER', 'DECEMBER']
            month_dtype = pd.api.types.CategoricalDtype(categories=month_order, ordered=True)
            df['Month'] = df['Month'].astype(month_dtype)
            print(df)
            pivot_data = df.pivot_table( index="Month",columns="Year",values=["Gross","Nett"],aggfunc="sum",margins=True, margins_name='Total') 
            pivot_data = pivot_data.applymap(lambda x: "{:,.2f}".format(x))
            pivot_data = pivot_data.drop(columns=[('Nett',          'Total'),('Gross',          'Total')])
            
    except Exception as e:
        print(e,".................")
    try:
        if month:
          
            source_data = source_data.filter(Month__in=month)
            print(source_data)
            df = pd.DataFrame(source_data.values())
            month_order = ['JANUARY', 'FEBRUARY', 'MARCH', 'APRIL', 'MAY', 'JUNE', 'JULY', 'AUGUST', 'SEPTEMBER', 'OCTOBER', 'NOVEMBER', 'DECEMBER']
            month_dtype = pd.api.types.CategoricalDtype(categories=month_order, ordered=True)
            df['Month'] = df['Month'].astype(month_dtype)
           
            pivot_data = df.pivot_table( index="Month",columns="Year",values=["Gross","Nett"],aggfunc="sum",margins=True, margins_name='Total')
            pivot_data = pivot_data.applymap(lambda x: "{:,.2f}".format(x))
            pivot_data = pivot_data.drop(columns=[('Nett',          'Total'),('Gross',          'Total')])
    except Exception as e:
        print(e,"...................")
    try:
        if country:

           
            source_data = source_data.filter(Country= country)


            df = pd.DataFrame(source_data.values())
            month_order = ['JANUARY', 'FEBRUARY', 'MARCH', 'APRIL', 'MAY', 'JUNE', 'JULY', 'AUGUST', 'SEPTEMBER', 'OCTOBER', 'NOVEMBER', 'DECEMBER']
            month_dtype = pd.api.types.CategoricalDtype(categories=month_order, ordered=True)
            df['Month'] = df['Month'].astype(month_dtype)
          

            pivot_data = df.pivot_table( index="Month",columns="Year",values=["Gross","Nett"],aggfunc="sum",margins=True, margins_name='Total')
            pivot_data = pivot_data.applymap(lambda x: "{:,.2f}".format(x))
            pivot_data = pivot_data.drop(columns=[('Nett',          'Total'),('Gross',          'Total')])
    except Exception as e:
        print(e,".......................")
    # pivot_data = df[(df.Year == '2021')|(df.Year =='2022')|(df.Year == '2022_FORECAST')].pivot_table( index="Month",columns=["Year"],values=["Gross","Nett"],aggfunc="sum",sort=False)
   
    
    try:
        if airline:
          
            source_data = source_data.filter(Airline__in= airline)
            df = pd.DataFrame(source_data.values())
            month_order = ['JANUARY', 'FEBRUARY', 'MARCH', 'APRIL', 'MAY', 'JUNE', 'JULY', 'AUGUST', 'SEPTEMBER', 'OCTOBER', 'NOVEMBER', 'DECEMBER']
            month_dtype = pd.api.types.CategoricalDtype(categories=month_order, ordered=True)
            df['Month'] = df['Month'].astype(month_dtype)
          
            pivot_data = df.pivot_table( index="Month",columns=["Year"],values=["Gross","Nett"],aggfunc="sum",margins=True, margins_name='Total')
            pivot_data = pivot_data.applymap(lambda x: "{:,.2f}".format(x))
            pivot_data = pivot_data.drop(columns=[('Nett',          'Total'),('Gross',          'Total')])
           
            # pivot_data = pivot_data.to_html()
    except Exception as e:
        print(e,".......................")
    try:
        if IATA_numeric_code:
           
            source_data = source_data.filter(IATA_Numeric_Code__in= IATA_numeric_code)
            df = pd.DataFrame(source_data.values())
            month_order = ['JANUARY', 'FEBRUARY', 'MARCH', 'APRIL', 'MAY', 'JUNE', 'JULY', 'AUGUST', 'SEPTEMBER', 'OCTOBER', 'NOVEMBER', 'DECEMBER']
            month_dtype = pd.api.types.CategoricalDtype(categories=month_order, ordered=True)
            df['Month'] = df['Month'].astype(month_dtype)
          
            pivot_data = df.pivot_table( index="Month",columns="Year",values=["Gross","Nett"],aggfunc="sum",margins=True, margins_name='Total')
            pivot_data = pivot_data.applymap(lambda x: "{:,.2f}".format(x))
            pivot_data = pivot_data.drop(columns=[('Nett',          'Total'),('Gross',          'Total')])
          
            # pivot_data = pivot_data.to_html()
    except Exception as e:
        print(e,".....................")   

    try:
        if agreement_type:
           
            source_data = source_data.filter(Agreement_Type__in= agreement_type)
            df = pd.DataFrame(source_data.values())
            month_order = ['JANUARY', 'FEBRUARY', 'MARCH', 'APRIL', 'MAY', 'JUNE', 'JULY', 'AUGUST', 'SEPTEMBER', 'OCTOBER', 'NOVEMBER', 'DECEMBER']
            month_dtype = pd.api.types.CategoricalDtype(categories=month_order, ordered=True)
            df['Month'] = df['Month'].astype(month_dtype)
         
            pivot_data = df.pivot_table( index="Month",columns="Year",values=["Gross","Nett"],aggfunc="sum",margins=True, margins_name='Total')
            pivot_data = pivot_data.applymap(lambda x: "{:,.2f}".format(x))
            pivot_data = pivot_data.drop(columns=[('Nett',          'Total'),('Gross',          'Total')])
         
    except Exception as e:
         print(e,".........................")
     
    try:          
        if revenue:
            # df = pd.DataFrame(list(DashboardModel.objects.all().values()))
            df = pd.DataFrame(list(source_data.values()))
            month_order = ['JANUARY', 'FEBRUARY', 'MARCH', 'APRIL', 'MAY', 'JUNE', 'JULY', 'AUGUST', 'SEPTEMBER', 'OCTOBER', 'NOVEMBER', 'DECEMBER']
            month_dtype = pd.api.types.CategoricalDtype(categories=month_order, ordered=True)
            df['Month'] = df['Month'].astype(month_dtype)
            pivot_data = df.pivot_table( index="Month",columns=["Year"],values=[revenue],aggfunc="sum",margins=True, margins_name='Total')
            pivot_data = pivot_data.applymap(lambda x: "{:,.2f}".format(x))
            pivot_data = pivot_data.drop(columns=[('Nett',          'Total'),('Gross',          'Total')])
    except Exception as e:
        print(e,"............................")
    
    if not any([airline_IATA_code, revenue, agreement_type, IATA_numeric_code, airline, country, month, year, region]):
      
        df = pd.DataFrame(list(source_data.values()))
        month_order = ['JANUARY', 'FEBRUARY', 'MARCH', 'APRIL', 'MAY', 'JUNE', 'JULY', 'AUGUST', 'SEPTEMBER', 'OCTOBER', 'NOVEMBER', 'DECEMBER']
        month_dtype = pd.api.types.CategoricalDtype(categories=month_order, ordered=True)
        df['Month'] = df['Month'].astype(month_dtype)
        pivot_data = df[(df.Year == '2022')|(df.Year =='2023')|(df.Year == '2023_FORECASTE')].pivot_table( index="Month",columns=["Year"],values=["Gross","Nett"],aggfunc="sum",margins=True, margins_name='Total')
        pivot_data = pivot_data.applymap(lambda x: "{:,.2f}".format(x))
        pivot_data = pivot_data.drop(columns=[('Nett',          'Total'),('Gross',          'Total')])
        reset = True
    # pivot_data = df.pivot_table( index="Month",columns="Year",values=["Gross","Nett"],aggfunc="sum",sort=False)
    pivot_data_forexel = pivot_data
    pivot_data = pivot_data.to_html()
    return pivot_data, pivot_data_forexel, df, reset
        
class Pivot(View):
    def get(self,request): 
        # DashboardModel.objects.filter(Year="2022_FORECAST").delete()
        
        user_admin_flag = False
        currnt_usr = request.user
        grps = currnt_usr.groups.all()
        grp_name = [grp.name for grp in grps]
        global_user = ["aasokan@airlinepros.net","amathew@airlinepros.com","vbaby@airlinepros.net","mjanardanan@airlinepros.net",
                       "sudhak@airlinepros.net","dpurushothaman@airlinepros.net","zaboobacker@airlinepros.net","snanavati@airlinepros.com",
                       "lbarber@airlinepros.net","admin@abhilash.com"]
        if currnt_usr.email in global_user:
            user_admin_flag = True
             
        if currnt_usr.email in global_user:
          
            pd.set_option('display.float_format', '{:.2f}'.format)
            df = pd.DataFrame(list(DashboardModel.objects.all().values()))
      
            df.replace(np. nan,0)
            source_data = DashboardModel.objects.all()
            is_filter = request.GET.get("is_filter")
            country = request.GET.get('country')
            flag_path = None
            if is_filter:
               
                airline_IATA_code = [i.strip() for i in json.loads(request.GET.get('airline_IATA_code'))] if request.GET.get('airline_IATA_code') else None

                country = request.GET.get('country')

                if country == "USA":
                    country_for_flag = "United States"
                else:
                    country_for_flag = country
                    
                if country_for_flag:
    
                    flag_path = str(Country.objects.get(name__iexact = country_for_flag).flag)



                region = request.GET.get('region')

                airline = [i.strip() for i in json.loads(request.GET.get('airline'))] if request.GET.get('airline') else None


                IATA_numeric_code = list(map(int,json.loads(request.GET.get('IATA_numeric_code')))) if request.GET.get('IATA_numeric_code') else None


                agreement_type = [i.strip() for i in json.loads(request.GET.get('agreement_type'))] if request.GET.get('agreement_type') else None



                revenue = request.GET.get("revenue")


                month = [i.strip() for i in json.loads(request.GET.get('month'))] if request.GET.get('month') else None


                year = [i.strip() for i in json.loads(request.GET.get('year'))] if request.GET.get('year') else None


                region = request.GET.getlist('region')
                reset = request.GET.get('reset')
                response = {}
              


                pivot_data, pivot_raw_data, graph, reset= filter(request,airline_IATA_code, region, year, month, country, airline, IATA_numeric_code , agreement_type , revenue)
                none_value_flag = False

                month_order = ['JANUARY', 'FEBRUARY', 'MARCH', 'APRIL', 'MAY', 'JUNE', 'JULY', 'AUGUST', 'SEPTEMBER', 'OCTOBER', 'NOVEMBER', 'DECEMBER']
                month_dtype = pd.api.types.CategoricalDtype(categories=month_order, ordered=True)
                df['Month'] = df['Month'].astype(month_dtype)

                try:
                    if reset:
                       
                        gross = graph[(graph.Year == '2022')|(graph.Year =='2023')|(graph.Year == '2023_FORECASTE')].pivot_table( index="Month",columns=["Year"],values=["Gross"],aggfunc="sum")
                        nett = graph[(graph.Year == '2022')|(graph.Year =='2023')|(graph.Year == '2023_FORECASTE')].pivot_table( index="Month",columns=["Year"],values=["Nett"],aggfunc="sum")
                    else:

                        gross = graph.pivot_table( index="Month",columns=["Year"],values=["Gross"],aggfunc="sum")
                        nett = graph.pivot_table( index="Month",columns=["Year"],values=["Nett"],aggfunc="sum")
                        

                    label = list(gross.index)
                
                    gross_year = MultiIndex.from_tuples(gross.columns)
                    values = gross.values
                    Gross_values = {}
                   
                    for i, years in enumerate(gross_year.get_level_values(1)):
                        Gross_values[years] = [row[i] for row in values]

                    nett_year = MultiIndex.from_tuples(nett.columns)
                    values = nett.values
                    Nett_values = {}
                    for i, years in enumerate(nett_year.get_level_values(1)):
                        Nett_values[years] = [row[i] for row in values]

                    try:
                        del request.session['last_filtered_pivot']
                    except:
                        pass
                    # try:
                    request.session['last_filtered_pivot'] = pivot_raw_data.to_json()
                    unique_country = DashboardModel.objects.all().values_list('Country', flat=True).distinct().order_by('Country')
                    unique_region = DashboardModel.objects.all().values_list('Region',flat=True).distinct().order_by('Region')

    
                    if country:   
                        unique_airline = DashboardModel.objects.filter(Country= country).values_list('Airline', flat=True).distinct().order_by('Airline')
                        unique_IATA_numeric_code = DashboardModel.objects.filter(Country= country).values_list('IATA_Numeric_Code', flat=True).distinct().order_by('IATA_Numeric_Code')
                        unique_agreement_type = DashboardModel.objects.filter(Country= country).values_list('Agreement_Type', flat=True).distinct().order_by('Agreement_Type')
                        unique_airline_IATA_code = DashboardModel.objects.filter(Country= country).values_list('Airline_IATA_Code', flat=True).distinct().order_by('Airline_IATA_Code')
                        unique_month = DashboardModel.objects.filter(Country= country).values_list('Month',flat=True).distinct().order_by('Month')
                        unique_year = DashboardModel.objects.filter(Country= country).values_list('Year',flat=True).distinct().order_by('Year')
                        unique_region = DashboardModel.objects.filter(Country= country).values_list('Region',flat=True).distinct().order_by('Region')
                    else:   
                        unique_airline = DashboardModel.objects.all().values_list('Airline', flat=True).distinct().order_by('Airline')
                        unique_IATA_numeric_code = DashboardModel.objects.all().values_list('IATA_Numeric_Code', flat=True).distinct().order_by('IATA_Numeric_Code')
                        unique_agreement_type = DashboardModel.objects.all().values_list('Agreement_Type', flat=True).distinct().order_by('Agreement_Type')
                        unique_airline_IATA_code = DashboardModel.objects.all().values_list('Airline_IATA_Code', flat=True).distinct().order_by('Airline_IATA_Code')
                        unique_month = DashboardModel.objects.all().values_list('Month',flat=True).distinct().order_by('Month')
                        unique_year = DashboardModel.objects.all().values_list('Year',flat=True).distinct().order_by('Year')
                        unique_region = DashboardModel.objects.all().values_list('Region',flat=True).distinct().order_by('Region')

                    # except:
                    #     pass
                    # pivot_data = filter(self, airline_IATA_code =airline_IATA_code, region =region, year =year, month =month, country =country, airline =airline, IATA_numeric_code =IATA_numeric_code, agreement_type =agreement_type, revenue =revenue)
                    output = render_to_string('pivot_table.html',{'no_data':False,'pivot_data':pivot_data,'label':label,'Gross_values':Gross_values,'Nett_values':Nett_values})
                    response['status'] = True
                  
                    response['filter_to_replace'] = render_to_string('filter.html',{'revenue_dropdown':revenue,'region_dropdown':region,'agreement_type_dropdown':agreement_type,'IATA_numeric_code_dropdown':IATA_numeric_code,'airline_IATA_code_dropdown':airline_IATA_code,'airline_dropdown':airline,'year_dropdown':year,'month_dropdown':month,'region_dropdown':region,'country_dropdown':country,'unique_airline':unique_airline,'unique_country':unique_country,'unique_airline_IATA_code':unique_airline_IATA_code,'unique_agreement_type':unique_agreement_type,
                                                                                    'unique_IATA_numeric_code':unique_IATA_numeric_code,'unique_airline':unique_airline,'unique_month':unique_month,'unique_year':unique_year,'user_admin_flag':user_admin_flag})
                    response['html_to_replace'] = output
                    response['flag_path'] = flag_path
                    return HttpResponse(json.dumps(response), content_type="application/json")
                
                
                except:
                    unique_country = DashboardModel.objects.all().values_list('Country', flat=True).distinct().order_by('Country')
                    unique_region = DashboardModel.objects.all().values_list('Region',flat=True).distinct().order_by('Region')

    
                    if country:   
                        unique_airline = DashboardModel.objects.filter(Country= country).values_list('Airline', flat=True).distinct().order_by('Airline')
                        unique_IATA_numeric_code = DashboardModel.objects.filter(Country= country).values_list('IATA_Numeric_Code', flat=True).distinct().order_by('IATA_Numeric_Code')
                        unique_agreement_type = DashboardModel.objects.filter(Country= country).values_list('Agreement_Type', flat=True).distinct().order_by('Agreement_Type')
                        unique_airline_IATA_code = DashboardModel.objects.filter(Country= country).values_list('Airline_IATA_Code', flat=True).distinct().order_by('Airline_IATA_Code')
                        unique_month = DashboardModel.objects.filter(Country= country).values_list('Month',flat=True).distinct().order_by('Month')
                        unique_year = DashboardModel.objects.filter(Country= country).values_list('Year',flat=True).distinct().order_by('Year')
                        unique_region = DashboardModel.objects.filter(Country= country).values_list('Region',flat=True).distinct().order_by('Region')
                    else:   
                        unique_airline = DashboardModel.objects.all().values_list('Airline', flat=True).distinct().order_by('Airline')
                        unique_IATA_numeric_code = DashboardModel.objects.all().values_list('IATA_Numeric_Code', flat=True).distinct().order_by('IATA_Numeric_Code')
                        unique_agreement_type = DashboardModel.objects.all().values_list('Agreement_Type', flat=True).distinct().order_by('Agreement_Type')
                        unique_airline_IATA_code = DashboardModel.objects.all().values_list('Airline_IATA_Code', flat=True).distinct().order_by('Airline_IATA_Code')
                        unique_month = DashboardModel.objects.all().values_list('Month',flat=True).distinct().order_by('Month')
                        unique_year = DashboardModel.objects.all().values_list('Year',flat=True).distinct().order_by('Year')
                        unique_region = DashboardModel.objects.all().values_list('Region',flat=True).distinct().order_by('Region')
                    output = render_to_string('pivot_table.html',{'no_data':True,'pivot_data':None,'label':None,'Gross_values':None,'Nett_values':None})
                    response['status'] = True
                    response['filter_to_replace'] = render_to_string('filter.html',{'revenue_dropdown':revenue,'region_dropdown':region,'agreement_type_dropdown':agreement_type,'IATA_numeric_code_dropdown':IATA_numeric_code,'airline_IATA_code_dropdown':airline_IATA_code,'airline_dropdown':airline,'year_dropdown':year,'month_dropdown':month,'region_dropdown':region,'country_dropdown':country,'unique_airline':unique_airline,'unique_country':unique_country,'unique_airline_IATA_code':unique_airline_IATA_code,'unique_agreement_type':unique_agreement_type,
                                                                                    'unique_IATA_numeric_code':unique_IATA_numeric_code,'unique_airline':unique_airline,'unique_month':unique_month,'unique_year':unique_year})
                    response['html_to_replace'] = output
                    return HttpResponse(json.dumps(response), content_type="application/json")
                    




            month_order = ['JANUARY', 'FEBRUARY', 'MARCH', 'APRIL', 'MAY', 'JUNE', 'JULY', 'AUGUST', 'SEPTEMBER', 'OCTOBER', 'NOVEMBER', 'DECEMBER']
            month_dtype = pd.api.types.CategoricalDtype(categories=month_order, ordered=True)
            df['Month'] = df['Month'].astype(month_dtype)

            pivot_data = df[(df.Year == '2022')|(df.Year =='2023')|(df.Year == '2023_FORECASTE')].pivot_table( index="Month",columns=["Year"],values=["Gross","Nett"],aggfunc="sum",margins=True, margins_name='Total')

            pivot_data = pivot_data.applymap(lambda x: "{:,.2f}".format(x))
            pivot_data = pivot_data.drop(columns=[('Nett',          'Total'),('Gross',          'Total')])
            


            gross = df[(df.Year == '2022')|(df.Year =='2023')|(df.Year == '2023_FORECASTE')].pivot_table( index="Month",columns=["Year"],values=["Gross"],aggfunc="sum")
            nett = df[(df.Year == '2022')|(df.Year =='2023')|(df.Year == '2023_FORECASTE')].pivot_table( index="Month",columns=["Year"],values=["Nett"],aggfunc="sum")
            label = list(gross.index)
                
            gross_year = MultiIndex.from_tuples(gross.columns)
            values = gross.values
            Gross_values = {}
            for i, year in enumerate(gross_year.get_level_values(1)):
                Gross_values[year] = [row[i] for row in values]
                                      
            nett_year = MultiIndex.from_tuples(nett.columns)
            values = nett.values
            Nett_values = {}
            for i, year in enumerate(nett_year.get_level_values(1)):
                Nett_values[year] = [row[i] for row in values]
          
            
    
            pivot_data = pivot_data.to_html()
            unique_country = DashboardModel.objects.all().values_list('Country', flat=True).distinct().order_by('Country')
            unique_airline = DashboardModel.objects.all().values_list('Airline', flat=True).distinct().order_by('Airline')
            unique_IATA_numeric_code = DashboardModel.objects.all().values_list('IATA_Numeric_Code', flat=True).distinct().order_by('IATA_Numeric_Code')
            unique_agreement_type = DashboardModel.objects.all().values_list('Agreement_Type', flat=True).distinct().order_by('Agreement_Type')
            unique_orc_rate = DashboardModel.objects.all().values_list('ORC_Rate', flat=True).distinct().order_by('ORC_Rate')
            unique_orc_actual = DashboardModel.objects.all().values_list('ORC_Actual', flat=True).distinct().order_by('ORC_Actual')
            unique_month = DashboardModel.objects.all().values_list('Month',flat=True).distinct().order_by('Month')
            unique_year = DashboardModel.objects.all().values_list('Year',flat=True).distinct().order_by('Year')
            unique_region = DashboardModel.objects.all().values_list('Region',flat=True).distinct().order_by('Region')
            unique_airline_IATA_code = DashboardModel.objects.all().values_list('Airline_IATA_Code', flat=True).distinct().order_by('Airline_IATA_Code')
            return render(request, 'pivot.html',{'pivot_data':pivot_data,'unique_country':unique_country,'unique_airline':unique_airline,'unique_IATA_numeric_code':unique_IATA_numeric_code,
                                                'unique_agreement_type':unique_agreement_type,'unique_orc_rate':unique_orc_rate,
                                                'unique_orc_actual':unique_orc_actual,'unique_month':unique_month,'unique_year':unique_year,'unique_region':unique_region,'unique_airline_IATA_code':unique_airline_IATA_code,
                                                'label':label,'Gross_values':Gross_values,'Nett_values':Nett_values,'user_admin_flag':user_admin_flag})
        else:  
          
            choosen_country_ = request.session.get('country')        
            obj = Country.objects.get(id = choosen_country_) 
            if obj.name == "United States":
                country_var = "USA"
            else:
                country_var = obj.name


            # if obj.name == "USA" or obj.name == "United States":
            pd.set_option('display.float_format', '{:.2f}'.format)
            df = pd.DataFrame(list(DashboardModel.objects.all().filter(Country__iexact= country_var).values()))

            df.replace(np. nan,0)
            source_data = DashboardModel.objects.filter(Country__iexact= country_var)
            is_filter = request.GET.get("is_filter")
            country = request.GET.get('country')    
            if is_filter:

                airline_IATA_code = json.loads(request.GET.get('airline_IATA_code')) if request.GET.get('airline_IATA_code') else None

                country = request.GET.get('country')
                region = request.GET.get('region')

                airline = [i.strip() for i in json.loads(request.GET.get('airline'))] if request.GET.get('airline') else None


                IATA_numeric_code = list(map(int,json.loads(request.GET.get('IATA_numeric_code')))) if request.GET.get('IATA_numeric_code') else None


                agreement_type = [i.strip() for i in json.loads(request.GET.get('agreement_type'))] if request.GET.get('agreement_type') else None



                revenue = request.GET.get("revenue")


                month = [i.strip() for i in json.loads(request.GET.get('month'))] if request.GET.get('month') else None


                year = [i.strip() for i in json.loads(request.GET.get('year'))] if request.GET.get('year') else None


                region = request.GET.getlist('region')
                reset = request.GET.get('reset')
                response = {}
             


                pivot_data, pivot_raw_data, graph, reset= filter(request,airline_IATA_code, region, year, month, country, airline, IATA_numeric_code , agreement_type , revenue)

                month_order = ['JANUARY', 'FEBRUARY', 'MARCH', 'APRIL', 'MAY', 'JUNE', 'JULY', 'AUGUST', 'SEPTEMBER', 'OCTOBER', 'NOVEMBER', 'DECEMBER']
                month_dtype = pd.api.types.CategoricalDtype(categories=month_order, ordered=True)
                df['Month'] = df['Month'].astype(month_dtype)

                try:
                    if reset:
                      
                        gross = graph[(graph.Year == '2022')|(graph.Year =='2023')|(graph.Year == '2023_FORECASTE')].pivot_table( index="Month",columns=["Year"],values=["Gross"],aggfunc="sum")
                        nett = graph[(graph.Year == '2022')|(graph.Year =='2023')|(graph.Year == '2023_FORECASTE')].pivot_table( index="Month",columns=["Year"],values=["Nett"],aggfunc="sum")
                    else:
                      
                        gross = graph.pivot_table( index="Month",columns=["Year"],values=["Gross"],aggfunc="sum")
                        nett = graph.pivot_table( index="Month",columns=["Year"],values=["Nett"],aggfunc="sum")
                    label = list(gross.index)
                
                    gross_year = MultiIndex.from_tuples(gross.columns)
                    values = gross.values
                    Gross_values = {}
                    for i, years in enumerate(gross_year.get_level_values(1)):
                        Gross_values[years] = [row[i] for row in values]

                    nett_year = MultiIndex.from_tuples(nett.columns)
                    values = nett.values
                    Nett_values = {}
                    for i, years in enumerate(nett_year.get_level_values(1)):
                        Nett_values[years] = [row[i] for row in values]

                    try:
                        del request.session['last_filtered_pivot']
                    except:
                        pass
                    # try:
                    request.session['last_filtered_pivot'] = pivot_raw_data.to_json()
                    unique_country = DashboardModel.objects.all().values_list('Country', flat=True).distinct().order_by('Country')
                    unique_airline = DashboardModel.objects.filter(Country__iexact= country_var).values_list('Airline', flat=True).distinct().order_by('Airline')
                    unique_IATA_numeric_code = DashboardModel.objects.filter(Country__iexact= country_var).values_list('IATA_Numeric_Code', flat=True).distinct().order_by('IATA_Numeric_Code')
                    unique_agreement_type = DashboardModel.objects.filter(Country__iexact= country_var).values_list('Agreement_Type', flat=True).distinct().order_by('Agreement_Type')
                    unique_orc_rate = DashboardModel.objects.filter(Country__iexact= country_var).values_list('ORC_Rate', flat=True).distinct().order_by('ORC_Rate')
                    unique_orc_actual = DashboardModel.objects.filter(Country__iexact= country_var).values_list('ORC_Actual', flat=True).distinct().order_by('ORC_Actual')
                    unique_month = DashboardModel.objects.filter(Country__iexact= country_var).values_list('Month',flat=True).distinct()
                    unique_year = DashboardModel.objects.filter(Country__iexact= country_var).values_list('Year',flat=True).distinct()
                    unique_region = DashboardModel.objects.filter(Country__iexact= country_var).values_list('Region',flat=True).distinct()
                    unique_airline_IATA_code = DashboardModel.objects.filter(Country__iexact= country_var).values_list('Airline_IATA_Code', flat=True).distinct().order_by('Airline_IATA_Code')
                    output = render_to_string('pivot_table.html',{'no_data':False,'pivot_data':pivot_data,'label':label,'Gross_values':Gross_values,'Nett_values':Nett_values})
                    response['status'] = True
                 
                    response['filter_to_replace'] = render_to_string('filter.html',{'revenue_dropdown':revenue,'region_dropdown':region,'agreement_type_dropdown':agreement_type,'IATA_numeric_code_dropdown':IATA_numeric_code,'airline_IATA_code_dropdown':airline_IATA_code,'airline_dropdown':airline,'year_dropdown':year,'month_dropdown':month,'region_dropdown':region,'country_dropdown':country,'unique_airline':unique_airline,'unique_country':unique_country,'unique_airline_IATA_code':unique_airline_IATA_code,'unique_agreement_type':unique_agreement_type,
                                                                                    'unique_IATA_numeric_code':unique_IATA_numeric_code,'unique_airline':unique_airline,'unique_month':unique_month,'unique_year':unique_year,'user_admin_flag':user_admin_flag})
                    response['html_to_replace'] = output
                    return HttpResponse(json.dumps(response), content_type="application/json")
                except:
                    unique_country = DashboardModel.objects.all().values_list('Country', flat=True).distinct().order_by('Country')
                    unique_airline = DashboardModel.objects.filter(Country__iexact= country_var).values_list('Airline', flat=True).distinct().order_by('Airline')
                    unique_IATA_numeric_code = DashboardModel.objects.filter(Country__iexact= country_var).values_list('IATA_Numeric_Code', flat=True).distinct().order_by('IATA_Numeric_Code')
                    unique_agreement_type = DashboardModel.objects.filter(Country__iexact= country_var).values_list('Agreement_Type', flat=True).distinct().order_by('Agreement_Type')
                    unique_orc_rate = DashboardModel.objects.filter(Country__iexact= country_var).values_list('ORC_Rate', flat=True).distinct().order_by('ORC_Rate')
                    unique_orc_actual = DashboardModel.objects.filter(Country__iexact= country_var).values_list('ORC_Actual', flat=True).distinct().order_by('ORC_Actual')
                    unique_month = DashboardModel.objects.filter(Country__iexact= country_var).values_list('Month',flat=True).distinct()
                    unique_year = DashboardModel.objects.filter(Country__iexact= country_var).values_list('Year',flat=True).distinct()
                    unique_region = DashboardModel.objects.filter(Country__iexact= country_var).values_list('Region',flat=True).distinct()
                    unique_airline_IATA_code = DashboardModel.objects.filter(Country__iexact= country_var).values_list('Airline_IATA_Code', flat=True).distinct().order_by('Airline_IATA_Code')
                    output = render_to_string('pivot_table.html',{'no_data':True,'pivot_data':None,'label':None,'Gross_values':None,'Nett_values':None})
                    response['status'] = True
                  
                    response['filter_to_replace'] = render_to_string('filter.html',{'revenue_dropdown':revenue,'region_dropdown':region,'agreement_type_dropdown':agreement_type,'IATA_numeric_code_dropdown':IATA_numeric_code,'airline_IATA_code_dropdown':airline_IATA_code,'airline_dropdown':airline,'year_dropdown':year,'month_dropdown':month,'region_dropdown':region,'country_dropdown':country,'unique_airline':unique_airline,'unique_country':unique_country,'unique_airline_IATA_code':unique_airline_IATA_code,'unique_agreement_type':unique_agreement_type,
                                                                                    'unique_IATA_numeric_code':unique_IATA_numeric_code,'unique_airline':unique_airline,'unique_month':unique_month,'unique_year':unique_year,'user_admin_flag':user_admin_flag})
                    response['html_to_replace'] = output
                    return HttpResponse(json.dumps(response), content_type="application/json")


            month_order = ['JANUARY', 'FEBRUARY', 'MARCH', 'APRIL', 'MAY', 'JUNE', 'JULY', 'AUGUST', 'SEPTEMBER', 'OCTOBER', 'NOVEMBER', 'DECEMBER']
            month_dtype = pd.api.types.CategoricalDtype(categories=month_order, ordered=True)
            df['Month'] = df['Month'].astype(month_dtype)

            pivot_data = df[(df.Year == '2022')|(df.Year =='2023')|(df.Year == '2023_FORECASTE')].pivot_table( index="Month",columns=["Year"],values=["Gross","Nett"],aggfunc="sum",margins=True, margins_name='Total')
            pivot_data = pivot_data.applymap(lambda x: "{:,.2f}".format(x))
            pivot_data = pivot_data.drop(columns=[('Nett',          'Total'),('Gross',          'Total')])
            

            gross = df[(df.Year == '2022')|(df.Year =='2023')|(df.Year == '2023_FORECASTE')].pivot_table( index="Month",columns=["Year"],values=["Gross"],aggfunc="sum")
            nett = df[(df.Year == '2022')|(df.Year =='2023')|(df.Year == '2023_FORECASTE')].pivot_table( index="Month",columns=["Year"],values=["Nett"],aggfunc="sum")

            label = list(gross.index)
                
            gross_year = MultiIndex.from_tuples(gross.columns)
            values = gross.values
            Gross_values = {}
            for i, year in enumerate(gross_year.get_level_values(1)):
                Gross_values[year] = [row[i] for row in values]
                                      
            nett_year = MultiIndex.from_tuples(nett.columns)
            values = nett.values
            Nett_values = {}
            for i, year in enumerate(nett_year.get_level_values(1)):
                Nett_values[year] = [row[i] for row in values]


            pivot_data = pivot_data.to_html()
            unique_country = DashboardModel.objects.all().values_list('Country', flat=True).distinct().order_by('Country')
            unique_airline = DashboardModel.objects.filter(Country__iexact= country_var).values_list('Airline', flat=True).distinct().order_by('Airline')
            unique_IATA_numeric_code = DashboardModel.objects.filter(Country__iexact= country_var).values_list('IATA_Numeric_Code', flat=True).distinct().order_by('IATA_Numeric_Code')
            unique_agreement_type = DashboardModel.objects.filter(Country__iexact= country_var).values_list('Agreement_Type', flat=True).distinct().order_by('Agreement_Type')
            unique_orc_rate = DashboardModel.objects.filter(Country__iexact= country_var).values_list('ORC_Rate', flat=True).distinct().order_by('ORC_Rate')
            unique_orc_actual = DashboardModel.objects.filter(Country__iexact= country_var).values_list('ORC_Actual', flat=True).distinct().order_by('ORC_Actual')
            unique_month = DashboardModel.objects.filter(Country__iexact= country_var).values_list('Month',flat=True).distinct()
            unique_year = DashboardModel.objects.filter(Country__iexact= country_var).values_list('Year',flat=True).distinct()
            unique_region = DashboardModel.objects.filter(Country__iexact= country_var).values_list('Region',flat=True).distinct()
            unique_airline_IATA_code = DashboardModel.objects.filter(Country__iexact= country_var).values_list('Airline_IATA_Code', flat=True).distinct().order_by('Airline_IATA_Code')
            return render(request, 'pivot.html',{'pivot_data':pivot_data,'unique_country':unique_country,'unique_airline':unique_airline,'unique_IATA_numeric_code':unique_IATA_numeric_code,
                                                'unique_agreement_type':unique_agreement_type,'unique_orc_rate':unique_orc_rate,
                                                'unique_orc_actual':unique_orc_actual,'unique_month':unique_month,'unique_year':unique_year,'unique_region':unique_region,'unique_airline_IATA_code':unique_airline_IATA_code,
                                                'label':label,'Gross_values':Gross_values,'Nett_values':Nett_values,'user_admin_flag':user_admin_flag})
      
   

class Getpivotdata(View):
    def get(self,request):
       
        choosen_country_ = request.session.get('country')
        obj = Country.objects.get(id = choosen_country_)
        
        response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

        if obj.name == "USA" or obj.name == "United States":
            response['Content-Disposition'] = 'attachment; filename=USA Data.xlsx'  
            my_pivot = request.session['last_filtered_pivot']
            my_pivot = pd.read_json(my_pivot )         
            my_pivot.to_excel("USA Data.xlsx")
            return response
                
        else:
            response['Content-Disposition'] = 'attachment; filename='+obj.name+" "+ 'Data.xlsx'
            my_pivot = request.session['last_filtered_pivot']
            my_pivot = pd.read_json(my_pivot )
            my_pivot.to_excel(obj.name+" "+'data.xlsx')

            return response
    
       

       
            
        
