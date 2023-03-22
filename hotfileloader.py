from ast import Break
import json
from tracemalloc import start
from unicodedata import name
from main.models import Airline, City, CommissionHistory, Country, State
from report.models import ReportPeriod, ReportFile, Charges, AgencyDebitMemo, Transaction, Taxes, Remittance, \
    DailyCreditCardFile, ReprocessFile, Disbursement, CarrierDeductions, Deduction
from agency.models import Agency
import django
import sys
import os
import re 
import datetime
from report.tasks import *
import logging


project_dir_path = os.path.abspath(os.getcwd())
sys.path.append(project_dir_path)
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'asplinks.settings')
django.setup()

ReportFile.objects.filter

def find_last_digit_value(digit):
    if digit == "}":
        return "-0"
    try:
        return int(digit)
    except:
        alpha =['{','A','B','C','D','E','F','G','H','I','}','J','K','L','M','N','O','P','Q','R']
        numeric= [0,1,2,3,4,5,6,7,8,9,-0,-1,-2,-3,-4,-5,-6,-7,-8,-9]
        dictval= dict(zip(alpha,numeric))
        return dictval[digit]



def amount_generator(amount, is_comm):
    amount= amount[:len(amount)-2]+"."+amount[len(amount)-2:]
    
    lastdgt= str(find_last_digit_value(amount[-1]))
    if '-' in lastdgt:
        lastdgt= lastdgt[1:]
        formatted_coblamt= amount.replace(amount[-1], str(lastdgt))
        for i in formatted_coblamt:
            if i != "0":
                amount_start_value= formatted_coblamt.index(i)
                break
        if not is_comm:
            amt= "-" + formatted_coblamt[amount_start_value:]
            return amt
        else:
            
            amt= formatted_coblamt[amount_start_value:]
    
    formatted_coblamt= amount.replace(amount[-1], str(lastdgt))
    for i in formatted_coblamt:
        if i != "0":
            amount_start_value= formatted_coblamt.index(i)
            break
    amt= formatted_coblamt[amount_start_value:]
    return amt


def rtncreate(agent_qs,rf,d):
    create_transaction(agent_qs,rf,d)


def process_billing_details_from_hotfile(text_file, request):
    
    
    user_country = request.session.get('country')
    user_country = Country.objects.get(pk=user_country)
    user_country_code= user_country.code
    content= open(text_file)


    content= content.readlines()

    aireline_qs= None
    


    for index in range(len(content)):
       
        txt = re.findall("BFH0000....01", content[index])

        if txt:
            txt= content[index]
            print(txt)
            bill_country_code= txt[36:36+2]
            airline_code= txt[16:16+3]
            if bill_country_code:
                print(bill_country_code,user_country_code,"--------------")
                if not bill_country_code.strip() == user_country_code.strip():
                    print("country not match...................")
                    return("country not match...please try other file....")
            if airline_code:
                try:
                    airline= Airline.objects.get(code= str(airline_code), country= user_country)
                    aireline_qs= airline
                except:
                    return("Airline not found belongs to the country... Please add.")
    ticket_found= False
    for index in range(len(content)):
        txt= re.findall("BKS0000....24", content[index])   
        if txt:
            ticket_found = True
            txt= content[index]
            print(txt,"........................")
            date= txt[13:13+6]
            break
    if not ticket_found:
        print("No tickets Found........")
        return ("No tickets.....")

    date= date
    print(date,"...........................raw")
    formatted_date= datetime.date( day= int(date[4:6]),month= int(date[2:4]),year= int("20"+date[0:2]))
    print(formatted_date,"................")

    all_reports= ReportPeriod.objects.filter(country= user_country)
    for i in range(len(all_reports)):
        start = all_reports[i].from_date
        end = all_reports[i].ped
        if start <= formatted_date <= end:
            pd= all_reports[i]
    ref_no= str(airline.code)+"-"+str(date[0:4])+"0"+str(pd.week)


    

    rf, created = ReportFile.objects.update_or_create(
            report_period=pd, airline=airline, country= user_country,
                 ref_no = ref_no, file= text_file.split('media/')[-1])

    


    #***************----------------------*********************




    starting_indexes=[]
    agent_index= []
    agency_id= []
    from agency.models import Agency

    for index in range(len(content)):
        txt= re.findall("BOH0000....03", content[index])
        agent_index.append(index) if txt else None
    for i in agent_index:
        txt= content[i]
        agent= txt[13:13+8]
        agency_id.append(agent)
    for i in range(len(content)):
        txt= re.findall("BKS0000....24", content[i])
        starting_indexes.append(i) if txt else None
    
    
    ticket_contents_list=[]

    last_ticket= len(starting_indexes)
    for ticket in range(len(starting_indexes)):
        if ticket== last_ticket-1:
            current_content= content[starting_indexes[ticket]:]
            ticket_contents_list.append(current_content)
            continue
        current_content=content[starting_indexes[ticket]:starting_indexes[ticket+1]]
        ticket_contents_list.append(current_content)

    li=[]


    for each_ticket in range(len(ticket_contents_list)):

        identifier= ticket_contents_list[each_ticket][0]
    
    
        for i in range(len(agent_index)):
            inde= content.index(identifier)
            if i != len(agent_index)-1:
                if inde >= agent_index[i] and inde <= agent_index[i+1]:
                    agent= agency_id[i]
                    break
            else:
                if inde >= agent_index[i]:
                    agent= agency_id[i]
                    break
                
        agent_qs_test= Agency.objects.filter(agency_no= agent)
        
        if agent_qs_test.exists():
            agent_qs= Agency.objects.get(agency_no= agent)
            
        else:
            agent_qs_test= Agency.objects.filter(agency_no= agent[:len(agent)-1])
            if agent_qs_test.exists():
                agent_qs= Agency.objects.get(agency_no= agent[:len(agent)-1])
        if not agent_qs_test.first():
            agent_qs= Agency.objects.create(agency_no = agent[:len(agent)-1], country = user_country)        
        
        

        tkt= ticket_contents_list[each_ticket]
        d={}
        tax=()
        isbks30_for_cobl= False
        fop= None
        
        TRNC= None
        ticket_set= set()
        date= None
        CUPI= None
        fare= None
        
        taxes= []
        taxtypes=[]
        rtvar= None
        transaction_amount_not_found = False
        t_a= None

        for j in tkt:
            if re.findall("BKS........45",j) and not rtvar:
                rft =j[25:25+14].strip()
                rft=rft[3:]
            
                rtvar= txt
                p={"transaction_type":"+RTDN", 'ticket_no':rft}

                rtncreate(agent_qs,rf,p)
            

            if re.findall("BKP........84",j):
                rtdn = j[46:46+19]
                
                if rtdn:
                    
                    rtdn= rtdn.strip()
                    rtdn= rtdn[3:13]
                    k={"transaction_type":"+RTDN", 'ticket_no':rtdn}
                    
                    

                    if rtdn.isdigit():
                        rtncreate(agent_qs,rf,k)
                        rtvar = rtdn
                        
                        #transaction ,b= create_transaction()
    #                     ##Transaction Create....
            
            if re.findall("BKS........30",j):

                type1 = j[62:62+8].strip()
                if type1:
                    amont1= amount_generator(j[70:70+11], False)
                    taxes.append(amont1)
                    taxtypes.append(type1)
                type2= j[81:81+8].strip()
                if type2:
                    amount2= amount_generator(j[89:89+11], False)
                    taxes.append(amount2)
                    taxtypes.append(type2)
                type3= j[100:100+8].strip()
                if type3:
                    amount3= amount_generator(j[108:108+11], False)
                    taxes.append(amount3)
                    taxtypes.append(type3)



            if re.findall('BKS........24', j):
            
            
                tkt_no=j[25:25+ 14].strip()
                tkt_no= tkt_no[len(tkt_no)-10:]
                ticket_set.add(tkt_no)
                #d['Ticket Number']= tkt_no
                # tkt_no = tkt_no[len(tkt_no)-10:]


                if not TRNC:
                    d['transaction_type']= j[71:71+4]
                    TRNC= j[71:71+4]

                if not date:
                    date= j[13:13+6].strip()
                    formatted_date= datetime.date( day= int(date[4:6]),month= int(date[2:4]),year= int("20"+date[0:2]))
                    # d['issue_date']= str(formatted_date)

                    date=str(formatted_date)
                    spdate= date.split('-')
                    spdate_month= datetime.datetime.strptime(str(spdate[1]), "%m")
                    spdate_month= spdate_month.strftime("%b")
                    spdate= spdate[-1]+ spdate_month +spdate[0][2:]
                    d['issue_date']= spdate
                if not CUPI:
                    d['cpui']= j[40:40+4].strip()
                    CUPI= j[40:40+4].strip()



            elif re.findall('BKS........39', j):
                d['stat']= j[40:40+3].strip()
                
                std_comm_rate= j[49:49+5]
            
                is_rate= False
                for i in std_comm_rate:
                    
                    if i!='0':
                        std_comm_rate= std_comm_rate[std_comm_rate.index(i):]
                        
                        std_comm_rate= std_comm_rate[:len(std_comm_rate)-2]+"."+std_comm_rate[len(std_comm_rate)-2:]
                        is_rate= True
                        d['std_comm_rate']= std_comm_rate
                        break
                if not is_rate:
                    d['std_comm_rate']= "00"
                        
                std_comm_amt= j[54:54+11]
                std_comm_amt= amount_generator(std_comm_amt, True)
                if TRNC==  'RFND':
                    
                    for m in tkt:
                        if re.findall("BKP........84",m):
                            tamount = m[39:46]
                            tamount = amount_generator(tamount,False)
                            t_a = tamount
                            if "-" in str(tamount):
                                d['transaction_amount']= t_a
                            else:
                                t_a = "-"+str(tamount)
                                d['transaction_amount']= t_a

                    d['std_comm_amount']= '-'+ str(std_comm_amt)
                else:
                    d['std_comm_amount']= str(std_comm_amt)





            elif re.findall('BKS........30',j):
            
                if not isbks30_for_cobl:
                    coblamt=j[40:40+11]
                    coblamt= amount_generator(coblamt, False)
                    d['cobl_amount']=coblamt
                    d['fare_amount']=coblamt
                    isbks30_for_cobl= True
                else:
                    continue

              

            elif re.findall("BAR........64",j):
                if not fare:
                    # f_a= re.sub('[^0-9.]'," ",str(j[53:53+12])).strip()
                    # if not f_a:
                    #     f_a= '0.0'
                    # d['fare_amount']=f_a
                    pass
                    fare=str(j[40:40+12])


                t_a = re.sub('[^0-9.]'," ",j[65:65+12]).strip()
                if not t_a:
                    t_a= "0.0"
                d['transaction_amount']= t_a
                if not t_a or t_a == "0.0":
                    transaction_amount_not_found = True
            elif re.findall("BKP........84",j):
                    if not fop:
                        d['fop']= j[25:25+2]
                        fop= j[25:25+2]
                    balance= j[97:97+11]
                    l= []
                    balance= amount_generator(balance, False)
                    l.append(balance)
                    d['balance']=balance      
                    if transaction_amount_not_found:
                        d['transaction_amount']= balance
                        t_a = balance

        if 'CP' in taxtypes:
            ind= taxtypes.index('CP')
            d['pen_type']= taxtypes[ind]
            d['pen']=taxes[ind]

        elif 'MF' in taxtypes:
            ind= taxtypes.index('MF')
            d['pen_type']= taxtypes[ind]
            d['pen_amount']=taxes[ind]
            
        if len(ticket_set)>1:
            a= list(ticket_set)

            tkt_no= a[0]+"-"+a[1][len(a[1])-2:]
            d['ticket_no']= tkt_no
            print(d['ticket_no'],".......")
        else:
            d['ticket_no']= list(ticket_set)[0]
            print(d['ticket_no'],".......")

        try:
            transaction ,b= create_transaction(agent_qs,rf,d)
        except Exception as e:
            print("creation failed...")
        for i in range(len(taxtypes)):        
            if taxtypes[i] in ['CP', 'MF', 'YQ', 'YR']:
                type_for_charges = taxtypes[i]
                amt_for_charges = taxes [i]
                charge = Charges.objects.update_or_create(amount= amt_for_charges, type= type_for_charges, transaction=transaction)

                print(amt_for_charges, type_for_charges,"charge created.........[][][[][][][][]", charge)
                type_for_charges= None
                amt_for_charges= None
            else:
                type_for_taxes = taxtypes[i]
                amt_for_taxes = taxes[i]
                tax = Taxes.objects.update_or_create(amount=amt_for_taxes, type=type_for_taxes, transaction=transaction)
                print(amt_for_taxes ,type_for_taxes, "taxes created.........[][][[][][][][]", tax)
                amt_for_taxes ,type_for_taxes= None, None 

        li.append(d)




        #************************------------------------------***********************

    return "Country incorrect"




def process_billing_details_from_hotfile_uk(text_file, request, pd):
    
    
    user_country = request.session.get('country')
    user_country = Country.objects.get(pk=user_country)
    user_country_code= user_country.code
    content= open(text_file)
    print(content,"............................content..................")

    content= content.readlines()

    aireline_qs= None
    


    for index in range(len(content)):
       
        txt = re.findall("BFH0000....01", content[index])

        if txt:
            txt= content[index]
            print(txt,"..........txt...///////////................")
            bill_country_code= txt[36:36+2]
            print(bill_country_code,"..............bill_country_code.............")
            airline_code= txt[16:16+3]
            print(airline_code,"................airline_code..............")
            if bill_country_code:
                print(bill_country_code,user_country_code,"--------------")
                if not bill_country_code.strip() == user_country_code.strip():
                    print("country not match...................")
                    return("country not match...please try other file....")
            if airline_code:
                try:
                    airline= Airline.objects.get(code= str(airline_code), country= user_country)
                  
                    aireline_qs= airline
                    print(aireline_qs,"...................aireline_qs...............")
                except:
                    return("Airline not found belongs to the country... Please add.")
    ticket_found= False
    for index in range(len(content)):
        txt= re.findall("BKS0000....24", content[index])   
        if txt:
            ticket_found = True
            txt= content[index]
            print(txt,"........................")
            date= txt[13:13+6]
            break
    if not ticket_found:
        print("No tickets Found........")
        return ("No tickets.....")

    date= date
    print(date,"...........................raw")
    formatted_date= datetime.date( day= int(date[4:6]),month= int(date[2:4]),year= int("20"+date[0:2]))
    print(formatted_date,"................")

    pd = ReportPeriod.objects.filter(country= user_country,year = int("20"+date[0:2]), month= int(date[2:4]) , week = int(pd[1])).first()

    # all_reports= ReportPeriod.objects.filter(country= user_country)
    # for i in range(len(all_reports)):
    #     start = all_reports[i].from_date
    #     end = all_reports[i].ped
    #     if start <= formatted_date <= end:
    #         pd= all_reports[i]
    ref_no= str(airline.code)+"-"+str(date[0:4])+"0"+str(pd.week)


    

    rf, created = ReportFile.objects.update_or_create(
            report_period=pd, airline=airline, country= user_country,
                 ref_no = ref_no, file= text_file.split('media/')[-1])

    


    #***************----------------------*********************




    starting_indexes=[]
    agent_index= []
    agency_id= []
    from agency.models import Agency

    for index in range(len(content)):
        txt= re.findall("BOH0000....03", content[index])
      
        agent_index.append(index) if txt else None
        print(agent_index,"................agent_index.................")
    for i in agent_index:
        txt= content[i]
        print(txt,"................txt...............")
        agent= txt[13:13+8]
        agency_id.append(agent)
        print(agency_id,"..............agency_id..................")
    for i in range(len(content)):
        txt= re.findall("BKS0000....24", content[i])
        starting_indexes.append(i) if txt else None
    
    
    ticket_contents_list=[]

    last_ticket= len(starting_indexes)
    for ticket in range(len(starting_indexes)):
        if ticket== last_ticket-1:
            current_content= content[starting_indexes[ticket]:]
            ticket_contents_list.append(current_content)
            continue
        current_content=content[starting_indexes[ticket]:starting_indexes[ticket+1]]
        ticket_contents_list.append(current_content)

    li=[]


    for each_ticket in range(len(ticket_contents_list)):

        identifier= ticket_contents_list[each_ticket][0]
    
    
        for i in range(len(agent_index)):
            inde= content.index(identifier)
            if i != len(agent_index)-1:
                if inde >= agent_index[i] and inde <= agent_index[i+1]:
                    agent= agency_id[i]
                    break
            else:
                if inde >= agent_index[i]:
                    agent= agency_id[i]
                    break
                
        agent_qs_test= Agency.objects.filter(agency_no= agent)
        
        if agent_qs_test.exists():
            agent_qs= Agency.objects.get(agency_no= agent)
            
        else:
            agent_qs_test= Agency.objects.filter(agency_no= agent[:len(agent)-1])
            if agent_qs_test.exists():
                agent_qs= Agency.objects.get(agency_no= agent[:len(agent)-1])
        if not agent_qs_test.first():
            agent_qs= Agency.objects.create(agency_no = agent[:len(agent)-1], country = user_country)        
        
        

        tkt= ticket_contents_list[each_ticket]
        d={}
        tax=()
        isbks30_for_cobl= False
        fop= None
        
        TRNC= None
        ticket_set= set()
        date= None
        CUPI= None
        fare= None
        
        taxes= []
        taxtypes=[]
        rtvar= None
        t_a = None



        for j in tkt:
            if re.findall("BKS........45",j) and not rtvar:
                rft =j[25:25+14].strip()
                rft=rft[3:]
            
                rtvar= txt
                p={"transaction_type":"+RTDN", 'ticket_no':rft}

                rtncreate(agent_qs,rf,p)
            

            if re.findall("BKP........84",j):
                rtdn = j[46:46+19]
                
                if rtdn:
                    
                    rtdn= rtdn.strip()
                    rtdn= rtdn[3:13]
                    k={"transaction_type":"+RTDN", 'ticket_no':rtdn}
                    
                    

                    if rtdn.isdigit():
                        rtncreate(agent_qs,rf,k)
                        rtvar = rtdn
                        
                        #transaction ,b= create_transaction()
    #                     ##Transaction Create....
            
            if re.findall("BKS........30",j):

                type1 = j[62:62+8].strip()
                if type1:
                    amont1= amount_generator(j[70:70+11], False)
                    taxes.append(amont1)
                    taxtypes.append(type1)
                type2= j[81:81+8].strip()
                if type2:
                    amount2= amount_generator(j[89:89+11], False)
                    taxes.append(amount2)
                    taxtypes.append(type2)
                type3= j[100:100+8].strip()
                if type3:
                    amount3= amount_generator(j[108:108+11], False)
                    taxes.append(amount3)
                    taxtypes.append(type3)



            if re.findall('BKS........24', j):
            
            
                tkt_no=j[25:25+ 14].strip()
                tkt_no= tkt_no[len(tkt_no)-10:]
                ticket_set.add(tkt_no)
                #d['Ticket Number']= tkt_no
                # tkt_no = tkt_no[len(tkt_no)-10:]


                if not TRNC:
                    d['transaction_type']= j[71:71+4]
                    TRNC= j[71:71+4]

                if not date:
                    date= j[13:13+6].strip()
                    formatted_date= datetime.date( day= int(date[4:6]),month= int(date[2:4]),year= int("20"+date[0:2]))
                    # d['issue_date']= str(formatted_date)

                    date=str(formatted_date)
                    spdate= date.split('-')
                    spdate_month= datetime.datetime.strptime(str(spdate[1]), "%m")
                    spdate_month= spdate_month.strftime("%b")
                    spdate= spdate[-1]+ spdate_month +spdate[0][2:]
                    d['issue_date']= spdate
                if not CUPI:
                    d['cpui']= j[40:40+4].strip()
                    CUPI= j[40:40+4].strip()



            elif re.findall('BKS........39', j):
                d['stat']= j[40:40+3].strip()
                
                std_comm_rate= j[49:49+5]
            
                is_rate= False
                for i in std_comm_rate:
                    
                    if i!='0':
                        std_comm_rate= std_comm_rate[std_comm_rate.index(i):]
                        
                        std_comm_rate= std_comm_rate[:len(std_comm_rate)-2]+"."+std_comm_rate[len(std_comm_rate)-2:]
                        is_rate= True
                        d['std_comm_rate']= std_comm_rate
                        break
                if not is_rate:
                    d['std_comm_rate']= "00"
                        
                std_comm_amt= j[54:54+11]
                std_comm_amt= amount_generator(std_comm_amt, True)
                if TRNC==  'RFND':
                    for m in tkt:
                        if re.findall("BKP........84",m):
                            tamount = m[39:46]
                            tamount = amount_generator(tamount,False)
                            t_a = tamount
                            if "-" in str(tamount):
                                d['transaction_amount']= t_a
                            else:
                                t_a = "-"+str(tamount)
                                d['transaction_amount']= t_a


                    d['std_comm_amount']= '-'+ str(std_comm_amt)
                else:
                    d['std_comm_amount']= str(std_comm_amt)





            elif re.findall('BKS........30',j):
            
                if not isbks30_for_cobl:
                    coblamt=j[40:40+11]
                    coblamt= amount_generator(coblamt, False)
                    d['cobl_amount']=coblamt
                    d['fare_amount']=coblamt
                    isbks30_for_cobl= True
                else:
                    continue

            elif re.findall("BKP........84",j):
                if not fop:
                    d['fop']= j[25:25+2]
                    fop= j[25:25+2]
                balance= j[97:97+11]
                l= []
                balance= amount_generator(balance, False)
                l.append(balance)
                d['balance']=balance            

            elif re.findall("BAR........64",j):
                if not fare:
                    # f_a= re.sub('[^0-9.]'," ",str(j[53:53+12])).strip()
                    # if not f_a:
                    #     f_a= '0.0'
                    # d['fare_amount']=f_a
                    pass
                    fare=str(j[40:40+12])
                if not t_a:

                    t_a = re.sub('[^0-9.]'," ",j[65:65+12]).strip()
                    if not t_a:
                        t_a= "0.0"
                    d['transaction_amount']= t_a


        if 'CP' in taxtypes:
            ind= taxtypes.index('CP')
            d['pen_type']= taxtypes[ind]
            d['pen']=taxes[ind]

        elif 'MF' in taxtypes:
            ind= taxtypes.index('MF')
            d['pen_type']= taxtypes[ind]
            d['pen_amount']=taxes[ind]
            
        if len(ticket_set)>1:
            a= list(ticket_set)

            tkt_no= a[0]+"-"+a[1][len(a[1])-2:]
            d['ticket_no']= tkt_no
            print(d['ticket_no'],".......")
        else:
            d['ticket_no']= list(ticket_set)[0]
            print(d['ticket_no'],".......")

        try:
            transaction ,b= create_transaction(agent_qs,rf,d)
        except Exception as e:
            print("creation failed...")
        for i in range(len(taxtypes)):        
            if taxtypes[i] in ['CP', 'MF', 'YQ', 'YR']:
                type_for_charges = taxtypes[i]
                amt_for_charges = taxes [i]
                charge = Charges.objects.update_or_create(amount= amt_for_charges, type= type_for_charges, transaction=transaction)

                print(amt_for_charges, type_for_charges,"charge created.........[][][[][][][][]", charge)
                type_for_charges= None
                amt_for_charges= None
            else:
                type_for_taxes = taxtypes[i]
                amt_for_taxes = taxes[i]
                tax = Taxes.objects.update_or_create(amount=amt_for_taxes, type=type_for_taxes, transaction=transaction)
                print(amt_for_taxes ,type_for_taxes, "taxes created.........[][][[][][][][]", tax)
                amt_for_taxes ,type_for_taxes= None, None 

        li.append(d)




        #************************------------------------------***********************

    return "Country incorrect"


