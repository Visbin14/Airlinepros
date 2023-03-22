from django import forms
from account.models import User
from django.contrib.auth.models import Group, Permission
from django.contrib.auth.forms import AuthenticationForm
from django.contrib.auth import (
    authenticate, get_user_model, password_validation,
)
from main.models import Country

from django_select2.forms import Select2MultipleWidget, Select2Widget


class UserForm(forms.ModelForm):
    """ Form used for user add/edit. """

    groups = forms.ModelMultipleChoiceField(queryset=Group.objects.order_by('name'), widget=Select2MultipleWidget, required=False)
    countries = forms.ModelMultipleChoiceField(queryset=Country.objects.order_by('name'), widget=Select2MultipleWidget, required=False)

    class Meta:
        model = User
        exclude = ('password','is_staff','date_joined','last_login','is_superuser')
        widgets = {
            'country': forms.Select(attrs={'class':'form-element-field form-control dropdown_arr_btn'})
        }

    def __init__(self, *args, **kwargs):
        super(UserForm, self).__init__(*args, **kwargs)
        if kwargs.get('instance'):
            if kwargs.get('instance').is_superuser:
                self.fields['groups'].required = False
                self.fields['groups'].widget.attrs.pop("autofocus", None)
                self.fields['countries'].widget.attrs.pop("autofocus", None)
        self.fields['countries'].initial = Country.objects.all()


class UserEditForm(forms.ModelForm):
    """ Form used for user add/edit. """

    groups = forms.ModelMultipleChoiceField(queryset=Group.objects.order_by('name'), widget=Select2MultipleWidget, required=False)
    countries = forms.ModelMultipleChoiceField(queryset=Country.objects.order_by('name'), widget=Select2MultipleWidget, required=False)

    class Meta:
        model = User
        exclude = ('email', 'password','is_staff','date_joined','last_login','is_superuser')
        widgets = {
            'country': forms.Select(attrs={'class':'form-element-field form-control dropdown_arr_btn'})
        }

    def __init__(self, *args, **kwargs):
        super(UserEditForm, self).__init__(*args, **kwargs)
        if kwargs.get('instance'):
            if kwargs.get('instance').is_superuser:
                self.fields['groups'].required = False
                self.fields['groups'].widget.attrs.pop("autofocus", None)
                self.fields['countries'].widget.attrs.pop("autofocus", None)


class UserProfileForm(forms.ModelForm):
    """ Form used for user add/edit. """

    class Meta:
        model = User
        exclude = ('email','password','is_staff','date_joined','last_login','is_superuser', 'groups', 'countries', 'is_active')


class GroupForm(forms.ModelForm):
    """ Form used for roles add/edit. """

    # permissions = forms.ModelMultipleChoiceField(queryset=Permission.objects.all(), widget=Select2MultipleWidget, required=False)

    class Meta:
        model = Group
        fields = ('name',)


class CustomAuthForm(AuthenticationForm):

    def clean(self):
        username = self.cleaned_data.get('username')
        password = self.cleaned_data.get('password')

        if username is not None and password:
            self.user_cache = authenticate(self.request, username=username, password=password)
            if self.user_cache is None:
                try:
                    user_temp = User.objects.get(email=username)
                except:
                    user_temp = None

                if user_temp is not None and user_temp.check_password(password):
                    self.confirm_login_allowed(user_temp)
                else:
                    raise forms.ValidationError(
                        self.error_messages['invalid_login'],
                        code='invalid_login',
                        params={'username': self.username_field.verbose_name},
                    )

        return self.cleaned_data
