import re
from django import forms
from django_select2.forms import Select2MultipleWidget

from agency.models import AgencyListReference, Agency, State, AgencyType, AgencyCollection
from account.models import User


def validate_agencies_no(value):
    if not re.match("^[0-9]+(,\s*[0-9]*)*$", value):
        raise forms.ValidationError(
            'Agency numbers not in valid format.')
    return value


def phone_number(value):
    if not re.match("^[+ Z0-9_ -]*$", value):
        raise forms.ValidationError("Not a valid phone number")
    return value


class AgencyListReferenceForm(forms.ModelForm):

    class Meta:
        model = AgencyListReference
        fields = ('file', 'file_type')


class AgencyTypeForm(forms.ModelForm):

    agencies = forms.CharField(widget=forms.Textarea,
                               validators=[validate_agencies_no], required=False,)

    class Meta:
        model = AgencyType
        fields = ('name', 'agencies')

    # def clean(self):
    #     cleaned_data = super(AgencyTypeForm, self).clean()
    #     name = cleaned_data.get('name')
    #     if self.instance.pk is not None:  # new instance only
    #         if self.instance.name != name:
    #             if AgencyType.objects.filter(name__iexact=name).exists():
    #                 self.add_error('name', 'AgencyType with that name already exists.')
    #     elif name and AgencyType.objects.filter(name__iexact=name).exists():
    #         self.add_error('name', 'AgencyType with that name already exists.')
    #     return cleaned_data


class AgencyCollectionForm(forms.ModelForm):
    agencies = forms.CharField(widget=forms.Textarea,
                               validators=[validate_agencies_no], required=False, )

    class Meta:
        model = AgencyCollection
        fields = ('name', 'agencies')



class AgencyForm(forms.ModelForm):
    # tel = forms.CharField(validators=[phone_number])

    class Meta:
        model = Agency
        fields = ("agency_no", "trade_name", "address1", "address2", "city", "state", "country",
                  "zip_code", "email", "vat_number", "tel", "agency_type", "home_agency", "sales_owner")

    def __init__(self, agency_types, cities, states, *args, **kwargs):
        super(AgencyForm, self).__init__(*args, **kwargs)
        self.fields['sales_owner'].queryset = User.objects.order_by('email')
        self.fields['state'].queryset = states
        self.fields['agency_type'].queryset = agency_types
        self.fields['city'].queryset = cities
        self.fields['city'].widget.attrs.update(
            {'class': 'dropdown_arr_btn form-element-field'})
        self.fields['state'].widget.attrs.update(
            {'class': 'dropdown_arr_btn form-element-field'})
        self.fields['country'].widget.attrs.update(
            {'class': 'dropdown_arr_btn form-element-field'})
        self.fields['agency_type'].widget.attrs.update(
            {'class': 'dropdown_arr_btn form-element-field'})
        self.fields['sales_owner'].widget.attrs.update(
            {'class': 'dropdown_arr_btn form-element-field'})



class StateOwnerForm(forms.ModelForm):

    # own_states = forms.ModelMultipleChoiceField(queryset=State.objects.all(), widget=Select2MultipleWidget,
    #                                              required=False, label='Your website')
    states = forms.ModelMultipleChoiceField(
        queryset=State.objects.order_by('name'), required=False)

    class Meta:
        model = User
        fields = ("own_states",)

    def __init__(self, user, request, *args, **kwargs):
        super(StateOwnerForm, self).__init__(*args, **kwargs)
        self.fields["states"].widget = Select2MultipleWidget()
        self.fields["states"].queryset = State.objects.filter(country=request.session.get('country')).order_by('name')
        self.fields['states'].required = False
        self.fields["states"].initial = State.objects.filter(
            owner=User.objects.get(pk=user)).values_list('id', flat=True)
