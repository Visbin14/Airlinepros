from datetime import datetime


def convert_amount(val):
    """
    183,939.59 to 183939.59
    """
    return val.replace(',', '') if val is not None else 0.0

def change_format(val):
    """
    98,543 to 98.543
    :param val:
    :return:
    """
    return str(val).replace(".","").replace(",", ".") if val is not None else "0.0"

def convert_transaction_date(val):
    """
    30JUL18  to a datetime object
    """
    # regex = re.compile(r'(?P<day>\d{2})(?P<month>\w{3})(?P<year>\d{2})')
    # match = re.match(regex, val)
    # if match:
    #     values = match.groupdict()
    return datetime.strptime(val, "%d%b%y").date() if val else None


def convert_date(val):
    """
    dd-MMM-YYYY date to a datetime object
    """
    return datetime.strptime(val, "%d-%b-%Y").date()


def get_agency_no(val):
    """
    61-5 0114 5 ==> 6150114
    """
    return val.replace(' ', '').replace('-', '')[:7]

def get_float(value):
    """
    converts the float value in the reports to a float value usable
    in python
    """
    if not value:
        return
    if value[value.__len__()-1] == '-':
        value = '-'+value[:value.__len__()-1]
    return float(value.replace(',',''))
