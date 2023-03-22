from django import forms

class Dashboardform(forms.Form):
    file = forms.FileField()
    