from django import forms
from .models import Airline
from main.models import Country
from account.models import User


class AirlineForm(forms.ModelForm):
    """ Airline Form. """

    class Meta:
        model = Airline
        exclude = ('arc_coordination_fee',)
        widgets = {
            'product_manager': forms.Select(attrs={'class': 'form-element-field form-control dropdown_arr_btn'})
        }

    def __init__(self, country, *args, **kwargs):
        super(AirlineForm, self).__init__(*args, **kwargs)
        self.fields['product_manager'].queryset = User.objects.order_by('email')
        if country:
            self.fields['country'].initial = Country.objects.get(pk=country)

    def clean_code(self):
        code = self.cleaned_data['code']
        if Airline.objects.filter(country=self.data['country'], code=self.cleaned_data['code']).exclude(
            id=self.instance.id).first():
            raise forms.ValidationError('Airline with this Country and code already exists.')
        return code

    def clean_abrev(self):
        abrev = self.cleaned_data['abrev']
        if Airline.objects.filter(country=self.data['country'], abrev=self.cleaned_data['abrev']).exclude(
            id=self.instance.id).first():
            raise forms.ValidationError('Airline with this Country and code already exists.')
        return abrev


class UsAirlineForm(forms.ModelForm):
    """US Airline Form. """

    class Meta:
        model = Airline
        exclude = ('arc_coordination_fee', 'accepts_AMEX', 'accepts_MC', 'accepts_VI', 'accepts_UATP')
        widgets = {
            'product_manager': forms.Select(attrs={'class': 'form-element-field form-control dropdown_arr_btn'})
        }

    def __init__(self, country, *args, **kwargs):
        super(UsAirlineForm, self).__init__(*args, **kwargs)
        self.fields['product_manager'].queryset = User.objects.order_by('email')
        if country:
            self.fields['country'].initial = Country.objects.get(pk=country)

    def clean_code(self):
        code = self.cleaned_data['code']
        if Airline.objects.filter(country=self.data['country'], code=self.cleaned_data['code']).exclude(
            id=self.instance.id).first():
            raise forms.ValidationError('Airline with this Country and code already exists.')
        return code

    def clean_abrev(self):
        abrev = self.cleaned_data['abrev']
        if Airline.objects.filter(country=self.data['country'], abrev=self.cleaned_data['abrev']).exclude(
            id=self.instance.id).first():
            raise forms.ValidationError('Airline with this Country and code already exists.')
        return abrev


class CountryForm(forms.ModelForm):
    """ Country Form. """

    class Meta:
        model = Country
        fields = '__all__'

    def clean_name(self):
        name = self.cleaned_data['name']
        if Country.objects.filter(name=name).exclude(
            id=self.instance.id).first():
            raise forms.ValidationError('Country with this name already exists.')
        return name

    def clean_code(self):
        code = self.cleaned_data['code']
        if Country.objects.filter(name=code).exclude(
            id=self.instance.id).first():
            raise forms.ValidationError('Country with this code already exists.')
        return code
