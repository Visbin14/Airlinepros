import calendar
from decimal import *
from django import template

register = template.Library()


@register.filter
def month_name(month_number):
    return calendar.month_name[month_number]


@register.filter
def add_decimal(num1, num2):
    return Decimal(num1) + Decimal(num2)



# @register.simple_tag
# def subtract_values(value1, value2, value3):
#     print("value1   ", value1, "    value2   ", value2, "    value3   ", value3)
#     # print("value1   ", type(value1), "    value2   ", type(value2), "    value3   ", type(value3))
#     # tt = value1 + value2 - value3
#     # print("---------------------        ", tt)
#     return value1 + value2 - value3

@register.simple_tag
def subtract_values(value1, value2, value3):

    print("value1   ", value1, "    value2   ", value2, "    value3   ", value3)
    if value1 is None or value1 == "":
        value1 = 0
    if value2 is None or value2 == "":
        value2 = 0
    if value3 is None or value3 == "":
        value3 = 0
    return value1 + value2 - value3

