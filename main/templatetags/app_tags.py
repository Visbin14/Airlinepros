from django import template

register = template.Library()


@register.filter
def get(dictn, key, default='0.0'):
    try:
        if not dictn.get(key):
            return 0.0
        return dictn.get(key, default)
    except:
        return default


@register.filter
def roundfloat(flt, places):
    try:
        if not isinstance(flt, float):
            return flt
        return round(flt, places)
    except:
        return flt