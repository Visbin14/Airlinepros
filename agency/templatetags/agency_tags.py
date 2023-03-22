import calendar
import datetime
from decimal import *
from django import template

register = template.Library()


@register.simple_tag
def week_number_of_month(date_value):
    return (date_value.isocalendar()[1] - date_value.replace(day=1).isocalendar()[1] + 1)


@register.simple_tag
def change_dateformat(date_value):
    return datetime.datetime.strftime(date_value,"%Y-%m-%d")
