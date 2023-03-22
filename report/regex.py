import re


file_header = re.compile(
    r'^\s+FCAIBILLDET\s+AIRLINE BILLING DETAILS\s+(?P<code>\d{3})\s+(?P<name>.*)$')

file_header_fr = re.compile(
    r'^\s+FCAIBILLDETFR\s+AIRLINE BILLING DETAILS FRANCE\s+(?P<code>\d{3})\s+(?P<name>.*)$')
date_pattern = '\d{2}\-[A-Z]{3}\-\d{4}'

period_header = re.compile(
    r'^Billing Period\:\d{6}\s+\((?P<start>' + date_pattern + ')\sto\s(?P<end>' + date_pattern + ') \)\s+REFERENCE:\s+(?P<ref_no>.*)$')
date_header = re.compile(
    r'^\s*INVOICE DATE\:\s+(?P<date>\d{2}\-[A-Z]{3}\-\d{4})')
summary_header = re.compile('^\s*SUMMARY$')
currencypattern = '-?\d*,?\d*,?\d+\.\d{2}'
grand_total = re.compile(
    r'^\s*GRAND\s+TOTAL\s+\(CAD\)\s+(?P<transaction_amount>-?\d*,?\d*,?\d+\.\d{2})\s+(?P<fare_amount>-?\d*,?\d*,?\d+\.\d{2})\s+(?P<tax>-?\d*,?\d*,?\d+\.\d{2})\s+(?P<fandc>-?\d*,?\d*,?\d+\.\d{2})\s+(?P<pen>-?\d*,?\d*,?\d+\.\d{2})\s+(?P<cobl_amount>-?\d*,?\d*,?\d+\.\d{2})\s+(?P<std_comm>-?\d*,?\d*,?\d+\.\d{2})\s+(?P<supp_comm>-?\d*,?\d*,?\d+\.\d{2})\s+(?P<tax_on_comm>-?\d*,?\d*,?\d+\.\d{2})\s+(?P<balance>-?\d*,?\d*,?\d+\.\d{2})$')
scope_combined = re.compile(r'^\s{0,1}SCOPE\s+COMBINED$')
scope_international = re.compile(r'^\s{0,1}SCOPE\s+INTERNATIONAL$')

agency_details = re.compile(
    r'^\s+(?P<agency_no>\d{2}\-\d\s\d{4}\s\d)\s+(?P<agency_no_short>\w{2})*\s+(?P<trade_name>(\w+\s)*.*)\s+')

transaction_details_exception = re.compile(
    r'^(?P<transaction_type>\w*)\s+(?P<ticket_no>\d+)\s+(?P<issue_date>\d{2}\w{3}\d{2})\s+(?P<cpui>\w*)\s+(?P<stat>\w*)\s+(?P<fop>\w*)\s+(?P<transaction_amount>-?\d*,?\d*,?\d+\.\d{2})\s+(?P<fare_amount>-?\d*,?\d*,?\d+\.\d{2})\s{8,45}((?P<pen>-?\d*,?\d*,?\d+\.\d{2})\s+(?P<pen_type>\w{2}))*\s+(?P<cobl_amount>-?\d*,?\d*,?\d+\.\d{2})\s+(?P<std_comm_rate>-?\d*,?\d*,?\d+\.\d{2})\s+(?P<std_comm_amount>-?\d*,?\d*,?\d+\.\d{2})\s+(?P<sup_comm_rate>-?\d*,?\d*,?\d+\.\d{2})\s+(?P<sup_comm_amount>-?\d*,?\d*,?\d+\.\d{2})\s+(?P<tax_on_comm>-?\d*,?\d*,?\d+\.\d{2})\s+(?P<balance>-?\d*,?\d*,?\d+\.\d{2})')


transaction_details_till_tax_amount = re.compile(r'^(?P<transaction_type>\w*)\s+(?P<ticket_no>\d+)\s+(?P<issue_date>\d{2}\w{3}\d{'
                                 r'2})\s+(?P<cpui>\w*)\s+(?P<stat>\w*\**)\s+(?P<fop>\w*)\s+('
                                 r'?P<transaction_amount>-?\d*,?\d*,?\d+\.\d{2})\*{0,1}\s+(?P<fare_amount>-?\d*,?\d*,'
                                 r'?\d+\.\d{2})\*{0,1}\s+(?P<tax_amount>-?\d*,?\d*,?\d+\.\d{2}){0,1}\*{0,'
                                 r'1}\s*(?P<tax_type>\w{2}){0,1}\s+(?P<fandc_amount>-?[\d,\.]+)\s+(?P<fandc_type>\w{2}){0,1}\s+(?P<pen>-?[\d,\.]+)\s+(?P<pen_type>\w{2})\s+(?P<cobl_amount>-?\d*,?\d*,?\d+\.\d{2})\*{0,'
                                 r'1}\s+(?P<std_comm_rate>-?\d*,?\d*,?\d+\.\d{2})\*{0,1}\s+(?P<std_comm_amount>-?\d*,'
                                 r'?\d*,?\d+\.\d{2})\*{0,1}\s+(?P<sup_comm_rate>-?\d*,?\d*,?\d+\.\d{2})\*{0,'
                                 r'1}\s+(?P<sup_comm_amount>-?\d*,?\d*,?\d+\.\d{2})\*{0,1}')

transaction_details_without_fandc_and_pen = re.compile(r'^(?P<transaction_type>\w*)\s+(?P<ticket_no>\d+)\s+(?P<issue_date>\d{2}\w{3}\d{'
                                 r'2})\s+(?P<cpui>\w*)\s+(?P<stat>\w*\**)\s+(?P<fop>\w*)\s+('
                                 r'?P<transaction_amount>-?\d*,?\d*,?\d+\.\d{2})\*{0,1}\s+(?P<fare_amount>-?\d*,?\d*,'
                                 r'?\d+\.\d{2})\*{0,1}\s+(?P<tax_amount>-?\d*,?\d*,?\d+\.\d{2}){0,1}\*{0,'
                                 r'1}\s*(?P<tax_type>\w{2}){0,1}\s+(?P<cobl_amount>-?\d*,?\d*,?\d+\.\d{2})\*{0,'
                                 r'1}\s+(?P<std_comm_rate>-?\d*,?\d*,?\d+\.\d{2})\*{0,1}\s+(?P<std_comm_amount>-?\d*,'
                                 r'?\d*,?\d+\.\d{2})\*{0,1}\s+(?P<sup_comm_rate>-?\d*,?\d*,?\d+\.\d{2})\*{0,'
                                 r'1}\s+(?P<sup_comm_amount>-?\d*,?\d*,?\d+\.\d{2})\*{0,1}')


transaction_details_exception_right1 = re.compile(
    r'^(?P<tax_amount>-?\d*,?\d*,?\d+\.\d{2})\*{0,1}\s+(?P<tax_type>\w{2}){0,1}\s+(?P<fandc_amount>-?\d*,?\d*,?\d+\.\d{2}){0,1}\*{0,1}'
    r'\s+(?P<fandc_type>\w{2}){0,1}')


transaction_details = re.compile(r'^(?P<transaction_type>\w*)\s+(?P<ticket_no>\d+)\s+(?P<issue_date>\d{2}\w{3}\d{'
                                 r'2})\s+(?P<cpui>\w*)\s+(?P<stat>\w*\**)\s+(?P<fop>\w*)\s+('
                                 r'?P<transaction_amount>-?\d*,?\d*,?\d+\.\d{2})\*{0,1}\s+(?P<fare_amount>-?\d*,?\d*,'
                                 r'?\d+\.\d{2})\*{0,1}\s+(?P<tax_amount>-?\d*,?\d*,?\d+\.\d{2}){0,1}\*{0,'
                                 r'1}\s*(?P<tax_type>\w{2}){0,1}\s+(?P<fandc_amount>-?[\d,\.]+)\s+(?P<fandc_type>\w{2}){0,1}\s+((?P<pen>-?[\d,\.]+)\s+(?P<pen_type>\w{2}))*\s+(?P<cobl_amount>-?\d*,?\d*,?\d+\.\d{2})\*{0,'
                                 r'1}\s+(?P<std_comm_rate>-?\d*,?\d*,?\d+\.\d{2})\*{0,1}\s+(?P<std_comm_amount>-?\d*,'
                                 r'?\d*,?\d+\.\d{2})\*{0,1}\s+(?P<sup_comm_rate>-?\d*,?\d*,?\d+\.\d{2})\*{0,'
                                 r'1}\s+(?P<sup_comm_amount>-?\d*,?\d*,?\d+\.\d{2})\*{0,1}\s+(?P<tax_on_comm>-?\d*,'
                                 r'?\d*,?\d+\.\d{2})\*{0,1}\s+(?P<balance>-?\d*,?\d*,?\d+\.\d{2})\*{0,1}')

transaction_details_zimb = re.compile(r'^(?P<transaction_type>\w*)\s+(?P<ticket_no>\d+)\s+(?P<issue_date>\d{2}\w{3}\d{'
                                 r'2})\s+(?P<cpui>\w*)\s+(?P<stat>\w*\**)\s+(?P<fop>\w*)\s+('
                                 r'?P<transaction_amount>-?\d*,?\d*,?\d+\.\d{2})\*{0,1}\s+(?P<fare_amount>-?\d*,?\d*,'
                                 r'?\d+\.\d{2})\*{0,1}\s+(?P<tax_amount>-?\d*,?\d*,?\d+\.\d{2}){0,1}\*{0,'
                                 r'1}\s*(?P<tax_type>\w{2}){0,1}\s+(?P<fandc_amount>-?[\d,\.]+)\s+(?P<fandc_type>\w{2}){0,1}\s+(?P<pen>-?[\d,\.]+)\s+(?P<pen_type>\w{2}){0,1}\s+(?P<cobl_amount>-?\d*,?\d*,?\d+\.\d{2})\*{0,'
                                 r'1}\s+(?P<std_comm_rate>-?\d*,?\d*,?\d+\.\d{2})\*{0,1}\s+(?P<std_comm_amount>-?\d*,'
                                 r'?\d*,?\d+\.\d{2})\*{0,1}\s+(?P<sup_comm_rate>-?\d*,?\d*,?\d+\.\d{2})\*{0,'
                                 r'1}\s+(?P<sup_comm_amount>-?\d*,?\d*,?\d+\.\d{2})\*{0,1}\s+(?P<tax_on_comm>-?\d*,'
                                 r'?\d*,?\d+\.\d{2})\*{0,1}\s+(?P<balance>-?\d*,?\d*,?\d+\.\d{2})\*{0,1}')

transaction_details1 = re.compile(r'^(?P<transaction_type>\w*)\s+(?P<ticket_no>\d+)\s+(?P<issue_date>\d{2}\w{3}\d{'
                                 r'2})\s+(?P<cpui>\w*)\s+(?P<stat>\w*\**)\s+(?P<fop>\w*)\s+(?P<transaction_amount>-?[\d,\.]+)\*{0,1}\s+'
                                  r'(?P<fare_amount>-?[\d,\.]+)\*{0,1}\s+(?P<tax_amount>-?[\d,\.]+){0,1}\*{0,'
                                 r'1}\s*(?P<tax_type>\w{2}){0,1}\s+(?P<fandc_amount>-?[\d,\.]+)\s+(?P<fandc_type>\w{2}){0,1}\s+'
                                  r'((?P<pen>-?[\d,\.]+)\s+(?P<pen_type>\w{2}))*\s+(?P<cobl_amount>-?[\d,\.]+)\*{0,'
                                 r'1}\s+(?P<std_comm_rate>-?[\d,\.]+)\*{0,1}\s+(?P<std_comm_amount>-?[\d,\.]+)\*{0,1}\s+(?P<sup_comm_rate>-?[\d,\.]+)\*{0,'
                                 r'1}\s+(?P<sup_comm_amount>-?[\d,\.]+)\*{0,1}\s+(?P<tax_on_comm>-?[\d,\.]+)\*{0,1}\s+(?P<balance>-?[\d,\.]+)\*{0,1}')


transaction_details_without_pen = re.compile(r'^(?P<transaction_type>\w*)\s+(?P<ticket_no>\d+)\s+(?P<issue_date>\d{2}\w{3}\d{'
                                 r'2})\s+(?P<cpui>\w*)\s+(?P<stat>\w*)\s+(?P<fop>\w*)\s+('
                                 r'?P<transaction_amount>-?\d*,?\d*,?\d+\.\d{2})\*{0,1}\s+(?P<fare_amount>-?\d*,?\d*,'
                                 r'?\d+\.\d{2})\*{0,1}\s+(?P<tax_amount>-?\d*,?\d*,?\d+\.\d{2}){0,1}\*{0,'
                                 r'1}\s*(?P<tax_type>\w{2}){0,1}\s+(?P<fandc_amount>-?[\d,\.]+)\s+(?P<fandc_type>\w{2}){0,1}\s+(?P<cobl_amount>-?\d*,?\d*,?\d+\.\d{2})\*{0,'
                                 r'1}\s+(?P<std_comm_rate>-?\d*,?\d*,?\d+\.\d{2})\*{0,1}\s+(?P<std_comm_amount>-?\d*,'
                                 r'?\d*,?\d+\.\d{2})\*{0,1}\s+(?P<sup_comm_rate>-?\d*,?\d*,?\d+\.\d{2})\*{0,'
                                 r'1}\s+(?P<sup_comm_amount>-?\d*,?\d*,?\d+\.\d{2})\*{0,1}\s+(?P<tax_on_comm>-?\d*,'
                                 r'?\d*,?\d+\.\d{2})\*{0,1}\s+(?P<balance>-?\d*,?\d*,?\d+\.\d{2})\*{0,1}')

only_pen_value = re.compile(r'^((?P<pen>-?\d*,?\d*,?\d+\.\d{2})\*{0,'
                                 r'1}\s+(?P<pen_type>\w{2}))*\s')

transaction_details_without_tax_to_pen = re.compile(r'^(?P<transaction_type>\w*)\s+(?P<ticket_no>\d+)\s+(?P<issue_date>\d{2}\w{3}\d{'
                                 r'2})\s+(?P<cpui>\w*)\s+(?P<stat>\w*\**)\s+(?P<fop>\w*)\s+('
                                 r'?P<transaction_amount>-?\d*,?\d*,?\d+\.\d{2})\*{0,1}\s+(?P<fare_amount>-?\d*,?\d*,'
                                 r'?\d+\.\d{2})\*{0,1}\s+(?P<cobl_amount>-?\d*,?\d*,?\d+\.\d{2})\*{0,'
                                 r'1}\s+(?P<std_comm_rate>-?\d*,?\d*,?\d+\.\d{2})\*{0,1}\s+(?P<std_comm_amount>-?\d*,'
                                 r'?\d*,?\d+\.\d{2})\*{0,1}\s+(?P<sup_comm_rate>-?\d*,?\d*,?\d+\.\d{2})\*{0,'
                                 r'1}\s+(?P<sup_comm_amount>-?\d*,?\d*,?\d+\.\d{2})\*{0,1}\s+(?P<tax_on_comm>-?\d*,'
                                 r'?\d*,?\d+\.\d{2})\*{0,1}\s+(?P<balance>-?\d*,?\d*,?\d+\.\d{2})\*{0,1}')


transaction_details__without_tax_and_fandc = re.compile(r'^(?P<transaction_type>\w*)\s+(?P<ticket_no>\d+)\s+(?P<issue_date>\d{2}\w{3}\d{'
                                 r'2})\s+(?P<cpui>\w*)\s+(?P<stat>\w*\**)\s+(?P<fop>\w*)\s+('
                                 r'?P<transaction_amount>-?\d*,?\d*,?\d+\.\d{2})\*{0,1}\s+(?P<fare_amount>-?\d*,?\d*,'
                                 r'?\d+\.\d{2})\*{0,1}\s+((?P<pen>-?\d*,?\d*,?\d+\.\d{2})\*{0,'
                                 r'1}\s+(?P<pen_type>\w{2}))*\s+(?P<cobl_amount>-?\d*,?\d*,?\d+\.\d{2})\*{0,'
                                 r'1}\s+(?P<std_comm_rate>-?\d*,?\d*,?\d+\.\d{2})\*{0,1}\s+(?P<std_comm_amount>-?\d*,'
                                 r'?\d*,?\d+\.\d{2})\*{0,1}\s+(?P<sup_comm_rate>-?\d*,?\d*,?\d+\.\d{2})\*{0,'
                                 r'1}\s+(?P<sup_comm_amount>-?\d*,?\d*,?\d+\.\d{2})\*{0,1}\s+(?P<tax_on_comm>-?\d*,'
                                 r'?\d*,?\d+\.\d{2})\*{0,1}\s+(?P<balance>-?\d*,?\d*,?\d+\.\d{2})\*{0,1}')

transaction_details__without_fandc = re.compile(r'^(?P<transaction_type>\w*)\s+(?P<ticket_no>\d+)\s+(?P<issue_date>\d{2}\w{3}\d{'
                                 r'2})\s+(?P<cpui>\w*)\s+(?P<stat>\w*\**)\s+(?P<fop>\w*)\s+('
                                 r'?P<transaction_amount>-?\d*,?\d*,?\d+\.\d{2})\*{0,1}\s+(?P<fare_amount>-?\d*,?\d*,'
                                 r'?\d+\.\d{2})\*{0,1}\s+(?P<tax_amount>-?\d*,?\d*,?\d+\.\d{2}){0,1}\*{0,'
                                 r'1}\s*(?P<tax_type>\w{2}){0,1}\s+((?P<pen>-?\d*,?\d*,?\d+\.\d{2})\*{0,'
                                 r'1}\s+(?P<pen_type>\w{2}))*\s+(?P<cobl_amount>-?\d*,?\d*,?\d+\.\d{2})\*{0,'
                                 r'1}\s+(?P<std_comm_rate>-?\d*,?\d*,?\d+\.\d{2})\*{0,1}\s+(?P<std_comm_amount>-?\d*,'
                                 r'?\d*,?\d+\.\d{2})\*{0,1}\s+(?P<sup_comm_rate>-?\d*,?\d*,?\d+\.\d{2})\*{0,'
                                 r'1}\s+(?P<sup_comm_amount>-?\d*,?\d*,?\d+\.\d{2})\*{0,1}\s+(?P<tax_on_comm>-?\d*,'
                                 r'?\d*,?\d+\.\d{2})\*{0,1}\s+(?P<balance>-?\d*,?\d*,?\d+\.\d{2})\*{0,1}')


transaction_detailsSA1 = re.compile(r'^(?P<transaction_type>\w*)\s+(?P<ticket_no>\d+)\s+(?P<issue_date>\d{2}\w{3}\d{'
                                    r'2})\s+(?P<cpui>\w*)\s+(?P<stat>\w*\**)\s+(?P<fop>\w*)\s+('
                                    r'?P<transaction_amount>-?\d*,?\d*,?\d+\.\d{2})\*{0,1}\s+(?P<fare_amount>-?\d*,'
                                    r'?\d*,?\d+\.\d{2})\*{0,1}\s+((?P<tax_amount>-?\d*,?\d*,?\d+\.\d{2})\s)+(?P<tax_type>\w{2}){0,1}\s+(?P<fandc_amount>-?\d*,?\d*,?\d+\.\d{2}){0,'
                                    r'1}\*{0,1}\s*(?P<fandc_type>\w{2}){0,1}\s+((?P<pen>-\s*|\d*,?\d*,?\d+\.\d{2})\*{0,'
                                    r'1}\s+(?P<pen_type>\\s*|\w{2}))*\s+(?P<cobl_amount>-?\d*,?\d*,?\d+\.\d{2})\*{0,'
                                    r'1}\s+(?P<std_comm_rate>-?\d*,?\d*,?\d+\.\d{2})\*{0,'
                                    r'1}\s+(?P<std_comm_amount>-?\d*,?\d*,?\d+\.\d{2})\*{0,1}')


transaction_detailsSA11 = re.compile(
    r'^(?P<transaction_type>\w*)\s+(?P<ticket_no>\d+)\s+(?P<issue_date>\d{2}\w{3}\d{2})\s+(?P<cpui>\w*)\s+(?P<stat>\w*\**)\s+(?P<fop>\w*)\s+(?P<transaction_amount>-?\d*,?\d*,?\d+\.\d{2})\*{0,1}\s+(?P<pen_type>-?\d*,?\d*,?\d+\.\d{2})\*{0,1}\s+(?P<cobl_amount>-?\d*,?\d*,?\d+\.\d{2}){0,1}\*{0,1}')

transaction_details__no_transaction_amount = re.compile(
    r'^(?P<transaction_type>\w*)\s+(?P<ticket_no>\d+)\s+(?P<issue_date>\d{2}\w{3}\d{2})\s+(?P<cpui>\w*)\s+(?P<stat>\w*)\s+(?P<fare_amount>-?\d*,?\d*,?\d+\.\d{2})\s+(?P<tax_amount>-?\d*,?\d*,?\d+\.\d{2}){0,1}\s*(?P<tax_type>\w{2}){0,1}\s+(?P<fandc_amount>-?\d*,?\d*,?\d+\.\d{2}){0,1}\s*(?P<fandc_type>\w{2}){0,1}\s+((?P<pen>-?\d*,?\d*,?\d+\.\d{2})\s+(?P<pen_type>\w{2}))*\s+(?P<cobl_amount>-?\d*,?\d*,?\d+\.\d{2})\s+(?P<std_comm_rate>-?\d*,?\d*,?\d+\.\d{2})\s+(?P<std_comm_amount>-?\d*,?\d*,?\d+\.\d{2})\s+(?P<sup_comm_rate>-?\d*,?\d*,?\d+\.\d{2})\s+(?P<sup_comm_amount>-?\d*,?\d*,?\d+\.\d{2})\s+(?P<tax_on_comm>-?\d*,?\d*,?\d+\.\d{2})\s+(?P<balance>-?\d*,?\d*,?\d+\.\d{2})')
transaction_details__nr_code = re.compile(
    r'^(?P<transaction_type>\w*)\s+(?P<ticket_no>\d+)\s+(?P<issue_date>\d{2}\w{3}\d{2})\s+(?P<cpui>\w*)\s+(\w*:\w*)\s+(?P<stat>\w*\**)\s+(?P<fop>\w*)\s+(?P<transaction_amount>-?\d*,?\d*,?\d+\.\d{2})\*{0,1}\s+(?P<fare_amount>-?\d*,?\d*,?\d+\.\d{2})\*{0,1}\s+(?P<tax_amount>-?\d*,?\d*,?\d+\.\d{2}){0,1}\*{0,1}\s*(?P<tax_type>\w{2}){0,1}\s+(?P<fandc_amount>-?\d*,?\d*,?\d+\.\d{2}){0,1}\*{0,1}\s*(?P<fandc_type>\w{2}){0,1}\s+((?P<pen>-?\d*,?\d*,?\d+\.\d{2})\s+(?P<pen_type>\w{2}))*\s+(?P<cobl_amount>-?\d*,?\d*,?\d+\.\d{2})\*{0,1}\s+(?P<std_comm_rate>-?\d*,?\d*,?\d+\.\d{2})\*{0,1}\s+(?P<std_comm_amount>-?\d*,?\d*,?\d+\.\d{2})\*{0,1}\s+(?P<sup_comm_rate>-?\d*,?\d*,?\d+\.\d{2})\*{0,1}\s+(?P<sup_comm_amount>-?\d*,?\d*,?\d+\.\d{2})\*{0,1}\s+(?P<tax_on_comm>-?\d*,?\d*,?\d+\.\d{2})\*{0,1}\s+(?P<balance>-?\d*,?\d*,?\d+\.\d{2})\*{0,1}')
transaction_details_rtdn = re.compile(
    r'^(?P<transaction_type>\+\w*):\s+(?P<ticket_no>\d+)\s+(?P<cpui>\d+)\s+(?P<fop>\w+)*\s+(?P<transaction_amount>-?\d*,?\d*,?\d+\.\d{2})\*{0,1}')
transaction_details_sub = re.compile(
    r'^\s+(?P<tax_amount>-?\d*,?\d*,?\d+\.\d{2})\*{0,1}\s*(?P<tax_type>\w+)\s*((?P<fandc_amount>-?\d*,?\d*,?\d+\.\d{2})\*{0,1}\s*(?P<fandc_type>\w+))*')

transaction_details_exception_left1 = re.compile(
    r'^(?P<transaction_type>\w*)\s+(?P<ticket_no>\d+)\s+(?P<issue_date>\d{2}\w{3}\d{2})\s+(?P<cpui>\w*)\s+(?P<stat>\w*)\s+(?P<fop>\w*)\s+(?P<transaction_amount>-?\d*,?\d*,?\d+\.\d{2})\*{0,1}')

transaction_details_exception_left = re.compile(
    r'^(?P<transaction_type>\w*)\s+(?P<ticket_no>\d+)\s+(?P<issue_date>\d{2}\w{3}\d{2})\s+(?P<cpui>\w*)\s+(?P<stat>\w*)\s+(?P<fop>\w*)\s+(?P<transaction_amount>-?\d*,?\d*,?\d+\.\d{2})\*{0,1}')
transaction_details_exception_right = re.compile(
    r'^\s+\*{2}\s+(?P<transaction_amount>-?\d*,?\d*,?\d+\.\d{2})\*{0,1}\s+(?P<fare_amount>-?\d*,?\d*,?\d+\.\d{2})\*{0,1}\s+(?P<tax_amount>-?\d*,?\d*,?\d+\.\d{2}){0,1}\*{0,1}\s+(?P<tax_type>\w{2}){0,1}\s+(?P<fandc_amount>-?\d*,?\d*,?\d+\.\d{2}){0,1}\*{0,1}\s+(?P<fandc_type>\w{2}){0,1}\s+((?P<pen>-?\d*,?\d*,?\d+\.\d{2})\*{0,1}\s+(?P<pen_type>\w{2}))*\s+(?P<cobl_amount>-?\d*,?\d*,?\d+\.\d{2})\*{0,1}\s+(?P<std_comm_rate>-?\d*,?\d*,?\d+\.\d{2})\*{0,1}\s+(?P<std_comm_amount>-?\d*,?\d*,?\d+\.\d{2})\*{0,1}\s+(?P<sup_comm_rate>-?\d*,?\d*,?\d+\.\d{2})\*{0,1}\s+(?P<sup_comm_amount>-?\d*,?\d*,?\d+\.\d{2})\*{0,1}\s+(?P<tax_on_comm>-?\d*,?\d*,?\d+\.\d{2})\*{0,1}\s+(?P<balance>-?\d*,?\d*,?\d+\.\d{2})\*{0,1}')
transaction_details_exception_middle = re.compile(
    r'^\s+(?P<fop>\w+)\s+(?P<transaction_amount>-?\d*,?\d*,?\d+\.\d{2})\*{0,1}')
newline = re.compile(r'^(?P<newline>\s)$')
transaction_spdr = re.compile(
    r'^SPDR\s+(?P<ticket_no>\d+)\s+(?P<issue_date>\d{2}\w{3}\d{2})\s+(?P<stat>\w*\**)\s+(?P<fop>\w*)\s+(?P<transaction_amount>-?\d*,?\d*,?\d+\.\d{2})\*{0,1}\s+(?P<fare_amount>-?\d*,?\d*,?\d+\.\d{2})\*{0,1}\s+(?P<cobl_amount>-?\d*,?\d*,?\d+\.\d{2})\*{0,1}\s+(?P<std_comm_amount>-?\d*,?\d*,?\d+\.\d{2})\*{0,1}\s+(?P<sup_comm_amount>-?\d*,?\d*,?\d+\.\d{2})\*{0,1}\s+(?P<tax_on_comm>-?\d*,?\d*,?\d+\.\d{2})\*{0,1}\s+(?P<balance>-?\d*,?\d*,?\d+\.\d{2})\*{0,1}')


cann = re.compile(
    r'^(?P<transaction_type>CANN)\s+(?P<ticket_no>\d+)\s+(?P<issue_date>\d{2}\w{3}\d{2})\s+(?P<stat>\w*\**)\s+(?P<transaction_amount>-?\d*,?\d*,?\d+\.\d{2})\*{0,1}\s+(?P<std_comm_amount>-?\d*,?\d*,?\d+\.\d{2})\*{0,1}\s+(?P<sup_comm_amount>-?\d*,?\d*,?\d+\.\d{2})\*{0,1}\s+(?P<tax_on_comm>-?\d*,?\d*,?\d+\.\d{2})\*{0,1}')

c_file_header = re.compile('^\s+PCAIDLYDET\s+AIRLINE PAYMENT CARD DAILY DETAILS\s+(?P<code>\d{3})\s+(?P<name>.*)$')
c_card_details = re.compile('^\s*INVOICE NUMBER:\s+(?P<invoice_number>\w+)\s+(?P<card>(\w{1,}\/{0,1}\s{0,1}\w{0,})*)')
c_transaction = re.compile(
    '^(?P<agency_no>\d{2}\-\d\s\d{4}\s\d)\s+(?P<transaction_type>\w*)\s+(?P<ticket_no>\d+)\s+(?P<issue_date>\d{2}\w{3}\d{2})\s+(\w*)')


cr_header = re.compile(
    r'^RE?P(?:OR)?T ?ID - (?P<report_id>\w{3}\d{3}-\w)\s+AIRLINES REPORTING CORPORATION\s+REF NBR - (?P<ref_no>\d+[\d-]+)$')
cr_ped = re.compile(r'^PAGE\s+\d\s+\d+\s+CARRIER INVOICE\s+CUR PED\s+-\s+(?P<ped>\d{2}/\d{2}/\d{2})')
cr_airline = re.compile(r'^\s+(?P<code>\d+)-(?P<name>(\w+(\s\w*)*))')

cr_seperater = re.compile(r'\-+')


currency = '\d*,?\d*,?\d*\.?\d+-?'

disp_record = re.compile(
    "^[ ]+([457])[ ]+"  # index
    "(" + currency + ")$")  # amount

disp_record1 = re.compile(
    "^[ ]+([457])[ ]+"  # index
    "(" + currency + ")[ ]+"
    "(" + currency + ")$"
)

arc_tot = re.compile(
    "^[ ]+" + currency + "[ ]+" + currency + "[ ]+"
                                             "(" + currency + ")[ ]*(NA)?$")  # arc
arc_deduc = re.compile(
    "[ ]+ARC DEDUCTIONS[ ]+"
    "(" + currency + ")[ ]*(NA)?$")  # arc2
arc_fees = re.compile(
    "[ ]+ASP FEES[ ]+"
    "(" + currency + ")[ ]*(NA)?$")  # arc2
arc_rev = re.compile(
    "[ ]+ARC REVERSALS[ ]+"
    "(" + currency + ")[ ]*(NA)?$")  # arc2

arc_net = re.compile(
    "[ ]+NET DISBURSEMENT[ ]+"
    "(" + currency + ")[ ]*(NA)?$")  # arc2

transaction_detailsSA = re.compile(r'^(?P<transaction_type>\w*)\s+(?P<ticket_no>\d+)\s+(?P<issue_date>\d{2}\w{3}\d{'
                                   r'2})\s+(?P<cpui>\w*)\s+(?P<stat>\w*)\s+(?P<fop>\w*)\s+('
                                   r'?P<transaction_amount>-?\d*,?\d*,?\d+\.\d{2})\*{0,1}\s+(?P<fare_amount>-?\d*,'
                                   r'?\d*,?\d+\.\d{2})\*{0,1}\s+(?P<tax_amount>-\s*|\d*,?\d*,?\d+\.\d{2}){0,1}\*{0,'
                                   r'1}\s*(?P<tax_type>\s*|\w{2}){0,1}\s+(?P<fandc_amount>-\s*|\d*,?\d*,?\d+\.\d{2}){0,'
                                   r'1}\*{0,1}\s*(?P<fandc_type>\s*|\w{2}){0,1}\s+((?P<pen>-\s*|\d*,?\d*,?\d+\.\d{2})\*{0,'
                                   r'1}\s+(?P<pen_type>\\s*|\w{2}))*\s+(?P<cobl_amount>-?\d*,?\d*,?\d+\.\d{2})\*{0,'
                                   r'1}\s+(?P<std_comm_rate>-?\d*,?\d*,?\d+\.\d{2})\*{0,'
                                   r'1}\s+(?P<std_comm_amount>-?\d*,?\d*,?\d+\.\d{2})\*{0,1}')

transaction_detailsSA_less_data = re.compile(r'^(?P<transaction_type>\w*)\s+(?P<ticket_no>\d+)\s+(?P<issue_date>\d{2}\w{3}\d{'
                                   r'2})\s+(?P<stat>\w*\**)\s+(?P<fop>\w*)\s+('
                                   r'?P<transaction_amount>-?\d*,?\d*,?\d+\.\d{2})\*{0,1}\s+(?P<fare_amount>-?\d*,'
                                   r'?\d*,?\d+\.\d{2})\*{0,1}\s+(?P<tax_amount>-\s*|\d*,?\d*,?\d+\.\d{2}){0,1}\*{0,'
                                   r'1}\s*(?P<tax_type>\s*|\w{2}){0,1}\s+(?P<fandc_amount>-\s*|\d*,?\d*,?\d+\.\d{2}){0,'
                                   r'1}\*{0,1}\s*(?P<fandc_type>\s*|\w{2}){0,1}\s+((?P<pen>-\s*|\d*,?\d*,?\d+\.\d{2})\*{0,'
                                   r'1}\s+(?P<pen_type>\\s*|\w{2}))*\s+(?P<cobl_amount>-?\d*,?\d*,?\d+\.\d{2})\*{0,'
                                   r'1}\s+(?P<std_comm_rate>-?\d*,?\d*,?\d+\.\d{2})\*{0,'
                                   r'1}\s+(?P<std_comm_amount>-?\d*,?\d*,?\d+\.\d{2})\*{0,1}')

# EC Regex
transaction_details_guyana = re.compile(
    r'^(?P<transaction_type>\w*)\s+(?P<ticket_no>\d+)\s+(?P<issue_date>\d{2}\w{3}\d{'
    r'2})\s+(?P<cpui>\w*)\s+(?P<stat>\w*\**)\s+(?P<fop>\w*)\s+('
    r'?P<transaction_amount>-?\d*,?\d*)\s+(?P<fare_amount>-?\d*,?\d*)\*{0,'
    r'1}\s+(?P<tax_amount>-?\d*,?\d*){0,1}\*{0,1}\s+(?P<fandc_amount>-?\d*)'
    r'\s*(?P<pen>-?\d*)\s+(?P<cobl_amount>-?\d*,?\d*)\*{0,1}\s+(?P<std_comm_rate>-?\d.\d*)'
    r'\s+(?P<std_comm_amount>-?\d,?\d*)\s+(?P<sup_comm_rate>-?\d.\d*)\s+(?P<sup_comm_amount>-?\d.\d*)'
    r'\s+(?P<tax_on_comm>-?\d.\d*)\s+(?P<balance>-?\d*,\d*)')


transaction_detailsGY = re.compile(r'^(?P<transaction_type>\w*)\s+(?P<ticket_no>\d+)\s+(?P<issue_date>\d{2}\w{3}\d{'
                                   r'2})\s+(?P<cpui>\w*)\s+(?P<stat>\w*\**)\s+(?P<fop>\w*)\s+('
                                   r'?P<transaction_amount>-?\d*,?\d*)\s+(?P<fare_amount>-?\d*,?\d*)\*{0,'
                                   r'1}\s+(?P<tax_amount>-?\d*,?\d*){0,1}\*{0,1}\s+(?P<fandc_amount>-?\d*)\s*('
                                   r'?P<pen>-?\d*,?\d*)\*{0,1}\s+(?P<cobl_amount>-?\d*,?\d*)\*{0,'
                                   r'1}\s+(?P<std_comm_rate>-?\d.\d*)\s+(?P<std_comm_amount>-?\d*,?\d*)\*{0,'
                                   r'1}\s+(?P<sup_comm_rate>-?\d*.?\d*)\*{0,1}\s+(?P<sup_comm_amount>-?\d*.?\d*)\*{0,'
                                   r'1}\s+(?P<tax_on_comm>-?\d.\d*)\s')

transaction_details_ec = re.compile(r'^(?P<transaction_type>\w*)\s+(?P<ticket_no>\d+)\s+(?P<issue_date>\d{2}\w{3}\d{'
                                    r'2})\s+(?P<cpui>\w*)\s+(?P<stat>\w*\**)\s+(?P<fop>\w*)\s+('
                                    r'?P<transaction_amount>-?\d*,?\d*,?\d+\,\d{2})\*{0,1}\s+(?P<fare_amount>-?\d*,'
                                    r'?\d*,?\d+\,\d{2})\*{0,1}\s+(?P<tax_amount>-?\d*,?\d*,?\d+\,\d{2}){0,1}\*{0,'
                                    r'1}\s*(?P<tax_type>\w{2}){0,1}\s+(?P<fandc_amount>-?[\d,\.]+)\s+(?P<fandc_type>\w{2}){0,1}\s+((?P<pen>-?[\d,\.]+)\*{0,'
                                 r'1}\s+(?P<pen_type>\w{2}))*\s+(?P<cobl_amount>-?\d*,?\d*,?\d+\,\d{2})\*{0,'
                                    r'1}\s+(?P<std_comm_rate>-?\d*,?\d*,?\d+\,\d{2})\*{0,'
                                    r'1}\s+(?P<std_comm_amount>-?\d*,?\d*,?\d+\,\d{2})\*{0,1}\s+('
                                    r'?P<sup_comm_rate>-?\d*,?\d*,?\d+\,\d{2})\*{0,1}\s+(?P<sup_comm_amount>-?\d*,'
                                    r'?\d*,?\d+\,\d{2})\*{0,1}\s+(?P<tax_on_comm>-?\d*,?\d*,?\d+\,\d{2})\*{0,'
                                    r'1}\s+(?P<balance>-?\d*,?\d*,?\d+\,\d{2})\*{0,1}')

transaction_details_ec1 = re.compile(r'^(?P<transaction_type>\w*)\s+(?P<ticket_no>\d+)\s+(?P<issue_date>\d{2}\w{3}\d{'
                                    r'2})\s+(?P<cpui>\w*)\s+(?P<stat>\w*\**)\s+(?P<fop>\w*)\s+('
                                    r'?P<transaction_amount>-?[\d,\.]+)\s+(?P<fare_amount>-?[\d,\.]+)\s+(?P<tax_amount>-?\d*,?\d*,?\d+\,\d{2}){0,1}\*{0,'
                                    r'1}\s*(?P<tax_type>\w{2}){0,1}\s+(?P<cobl_amount>-?[\d,\.]+)\s+(?P<std_comm_rate>-?\d*,?\d*,?\d+\,\d{2})\*{0,'
                                    r'1}\s+(?P<std_comm_amount>-?\d*,?\d*,?\d+\,\d{2})\*{0,1}\s+('
                                    r'?P<sup_comm_rate>-?\d*,?\d*,?\d+\,\d{2})\*{0,1}\s+(?P<sup_comm_amount>-?\d*,'
                                    r'?\d*,?\d+\,\d{2})\*{0,1}\s+(?P<tax_on_comm>-?\d*,?\d*,?\d+\,\d{2})\*{0,'
                                    r'1}\s+(?P<balance>-?[\d,\.]+)')

transaction_details_exception_left_ec = re.compile(r'^(?P<transaction_type>\w*)\s+(?P<ticket_no>\d+)\s+(?P<issue_date>\d{2}\w{3}\d{'
                                    r'2})\s+(?P<cpui>\w*)\s+(?P<stat>\w*\**)\s+(?P<fop>\w*)\s+('
                                    r'?P<transaction_amount>-?[\d,\.]+)')


transaction_details_exception_right_ec = re.compile(
    r'^\s+(?P<fare_amount>-?[\d,\.]+)\s+(?P<tax_amount>-?\d*,?\d*,?\d+\,\d{2}){0,1}\*{0,'
                                    r'1}\s*(?P<tax_type>\w{2}){0,1}\s+(?P<cobl_amount>-?[\d,\.]+)\s+(?P<std_comm_rate>-?\d*,?\d*,?\d+\,\d{2})\*{0,'
                                    r'1}\s+(?P<std_comm_amount>-?\d*,?\d*,?\d+\,\d{2})\*{0,1}\s+('
                                    r'?P<sup_comm_rate>-?\d*,?\d*,?\d+\,\d{2})\*{0,1}\s+(?P<sup_comm_amount>-?\d*,'
                                    r'?\d*,?\d+\,\d{2})\*{0,1}\s+(?P<tax_on_comm>-?\d*,?\d*,?\d+\,\d{2})\*{0,'
                                    r'1}\s+(?P<balance>-?[\d,\.]+)')

transaction_details_sub_ec = re.compile(
    r'^\s+(?P<tax_amount>-?\d*,?\d*,?\d+\,\d{2})\*{0,1}\s*(?P<tax_type>\w{2})')


grand_total_ca_single_line_ec = re.compile(
    r'^\s{0,1}GRAND TOTAL\s+(?P<type>\w{0,2})\s+(?P<transaction_amount>-?[\d,\.]+)\s+(?P<fare_amount>-?[\d,\.]+)\s+(?P<tax>-?[\d,\.]+)+\s+(?P<fandc>-?[\d,\.]+)\s+(?P<pen>-?[\d,\.]+)\s+(?P<cobl_amount>-?[\d,\.]+)\s+(?P<std_comm>-?[\d,\.]+)\s+(?P<supp_comm>-?[\d,\.]+)\s+(?P<tax_on_comm>-?[\d,\.]+)\s+(?P<balance>-?[\d,\.]+)$')

