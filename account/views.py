from django.contrib import messages
from django.contrib.auth import login
from django.contrib.auth import views as auth_views
from django.contrib.auth.forms import PasswordResetForm
from django.contrib.auth.mixins import PermissionRequiredMixin
from django.contrib.auth.models import Group
from django.contrib.auth.models import Permission
from django.contrib.auth.tokens import PasswordResetTokenGenerator
from django.contrib.messages.views import SuccessMessageMixin
from django.core.exceptions import ValidationError
from django.db.models import Q, Value
from django.db.models.functions import Concat
from django.http import HttpResponse, HttpResponseRedirect
from django.shortcuts import get_object_or_404, redirect
from django.urls import reverse_lazy
from django.utils import six
from django.utils.http import urlsafe_base64_decode
from django.views import View
from django.views.generic import ListView, DetailView, CreateView, UpdateView, DeleteView

from account.models import User
from main.models import Country
from main.tasks import send_mail, is_arc
from .forms import UserForm, GroupForm, UserProfileForm, UserEditForm, CustomAuthForm

permissions_list = ['view_agency', 'change_agency', 'view_agencytype', 'change_agencytype', 'view_user', 'change_user',
                    'view_stateowners', 'change_stateowners', 'view_group', 'change_group', 'view_airline',
                    'change_airline', 'view_sales_details', 'download_sales_details', 'view_sales_summary','view_new_sales_summary',
                    'download_sales_summary','download_new_sales_summary', 'view_adm', 'download_adm', 'view_sales_by', 'download_sales_by',
                    'view_all_sales', 'download_all_sales', 'view_year_to_year', 'download_year_to_years',
                    'view_commission', 'download_commission', 'view_sales_comparison', 'download_sales_comparison',
                    'view_top_agency', 'download_top_agency', 'view_monthly_yoy', 'download_monthly_yoy',
                    'view_airline_agency', 'download_airline_agency', 'view_agencycollection',
                    'change_agencycollection', 'view_agency_collection_report', 'download_agency_collection_report',
                    'view_upload_reports', 'view_disbursement_summary', 'download_disbursement_summary',
                    'view_calendar', 'view_upload_calendar', 'view_airline_management', 'change_airline_management','view_weekly_sales_report','download_weekly_sales']


class TokenGenerator(PasswordResetTokenGenerator):
    """"Activation link token generator"""

    def _make_hash_value(self, user, timestamp):
        return (
            six.text_type(user.pk) + six.text_type(timestamp) +
            six.text_type(user.is_active)
        )


account_activation_token = TokenGenerator()


class EmailValidationOnForgotPassword(PasswordResetForm):
    """ Password Reset form. """

    def clean_email(self):
        email = self.cleaned_data['email']
        if not User.objects.filter(email__iexact=email, is_active=True).exists():
            raise ValidationError("This email id is not registered in this application")

        return email


class CustomLogin(auth_views.LoginView):
    """Login View."""

    redirect_authenticated_user = True
    form_class = CustomAuthForm

    def form_valid(self, form):
        login(self.request, form.get_user())
        if self.request.user.is_superuser:
            self.request.session['country'] = Country.objects.first().id
        else:
            if self.request.user.countries.first():
                self.request.session['country'] = self.request.user.countries.first().id
            else:
                self.request.session['country'] = Country.objects.first().id
        return HttpResponseRedirect(self.get_success_url())


class CustomPasswordReset(auth_views.PasswordResetView):
    """Password reset view."""

    html_email_template_name = 'registration/password_reset_email.html'
    form_class = EmailValidationOnForgotPassword


class UserListView(PermissionRequiredMixin, ListView):
    """Users listing with pagination."""

    model = User
    template_name = 'user-listing.html'
    context_object_name = 'users'
    permission_required = ('account.view_user',)

    def get_queryset(self):
        order = self.request.GET.get('order_by', 'email')
        status = self.request.GET.get('status', '')
        role = self.request.GET.get('role', '')
        qs = User.objects.prefetch_related('groups').annotate(
            name=Concat('first_name', Value(' '), 'last_name')).order_by(order)
        query = self.request.GET.get('q')
        if query:
            qs = qs.filter(
                Q(name__icontains=query) | Q(email__icontains=query)
            )
        if status:
            qs = qs.filter(is_active=status)
        if role:
            qs = qs.filter(groups__in=[role])

        return qs

    def get_context_data(self, **kwargs):
        context = super(UserListView, self).get_context_data(**kwargs)
        context['activate'] = 'users'
        context['order_by'] = self.request.GET.get('order_by', 'id')
        context['query'] = self.request.GET.get('q', '')
        context['status'] = self.request.GET.get('status', '')
        context['role'] = self.request.GET.get('role', '')
        context['groups'] = Group.objects.order_by('name')
        return context


class UserDetailsView(PermissionRequiredMixin, DetailView):
    """User details view."""

    model = User
    template_name = 'user-details.html'
    context_object_name = 'object'

    permission_required = ('account.view_user',)

    def get_context_data(self, **kwargs):
        context = super(UserDetailsView, self).get_context_data(**kwargs)
        return context


class UserCreateView(PermissionRequiredMixin, CreateView):
    """ User creation view."""

    model = User
    form_class = UserForm
    template_name = 'add-user.html'
    success_message = "%(name)s was created successfully"
    permission_required = ('account.change_user',)

    def form_valid(self, form):
        self.object = form.save()
        password = User.objects.make_random_password()
        self.object.set_password(password)
        self.object.save()
        context = {
            'user': self.object,
            'request': self.request,
            'password': password
        }
        send_mail("Account Created", "email/user-create-email.html", context, [self.object.email],
                  from_email='Assda@assda.com')
        messages.add_message(self.request, messages.SUCCESS,
                             'User added successfully')
        return HttpResponseRedirect(self.get_success_url())

    def get_context_data(self, **kwargs):
        context = super(UserCreateView, self).get_context_data(**kwargs)
        context['activate'] = 'users'
        return context


class UserUpdateView(PermissionRequiredMixin, UpdateView):
    """ User edit view."""

    form_class = UserForm
    model = User
    template_name = 'edit-user.html'

    context_object_name = 'object'
    permission_required = ('account.change_user',)

    def get_form_class(self):
        if self.request.user.pk == self.kwargs.get('pk'):
            return UserProfileForm
        else:
            return UserEditForm

    def get_context_data(self, **kwargs):
        context = super(UserUpdateView, self).get_context_data(**kwargs)
        context['activate'] = 'users'
        return context

    def form_valid(self, form):
        messages.add_message(self.request, messages.SUCCESS,
                             'User details updated successfully.')
        return super(UserUpdateView, self).form_valid(form)


# class UserDeleteView(PermissionRequiredMixin, DeleteView):
#     """User deletion view."""
#
#     model = User
#
#     success_url = '/users'
#     permission_required = ('account.change_user',)
#
#     def delete(self, *args, **kwargs):
#         self.get_object().delete()
#         return HttpResponseRedirect(self.success_url)
#
#     def test_func(self):
#         return True
#
#     def get_success_url(self):
#         return reverse_lazy('users')
#
#     def delete(self, request, *args, **kwargs):
#         messages.add_message(self.request, messages.SUCCESS,
#                              'User deleted successfully.')
#         return super(UserDeleteView, self).delete(request, *args, **kwargs)


class UserPasswordResetView(PermissionRequiredMixin, View):
    """User password reset view."""

    permission_required = ('account.change_user',)

    def post(self, request, pk):
        user = get_object_or_404(User, pk=pk)
        password = User.objects.make_random_password()
        user.set_password(password)
        user.save()
        form = PasswordResetForm(data={'email': user.email})
        form.is_valid()
        form.save(request=request,
                  html_email_template_name='registration/user_password_reset_email.html')
        return redirect('user_details', pk=user.id)


def activate(request, uidb64, token):
    """Activate account view."""
    try:
        uid = urlsafe_base64_decode(uidb64).decode()
        user = User.objects.get(pk=uid)
    except(TypeError, ValueError, OverflowError, User.DoesNotExist):
        user = None
    if user is not None and account_activation_token.check_token(user, token):
        user.is_active = True
        user.save()
        # login(request, user)
        messages.add_message(request, messages.SUCCESS, 'Account activated.')
        return redirect('login')
    else:
        return HttpResponse('Activation link is invalid!')


class RoleListView(PermissionRequiredMixin, ListView):
    """Roles listing with pagination."""

    model = Group
    template_name = 'role-listing.html'
    context_object_name = 'roles'
    permission_required = ('auth.view_group',)

    def get_queryset(self):
        order = self.request.GET.get('order_by', 'name')
        qs = Group.objects.order_by(order)
        query = self.request.GET.get('q')
        if query:
            qs = qs.filter(
                Q(name__icontains=query)
            )
        return qs

    def get_context_data(self, **kwargs):
        context = super(RoleListView, self).get_context_data(**kwargs)
        context['activate'] = 'users'
        context['order_by'] = self.request.GET.get('order_by', 'id')
        context['query'] = self.request.GET.get('q', '')
        return context


class RoleCreateView(PermissionRequiredMixin, SuccessMessageMixin, CreateView):
    """ Role creation view."""

    model = Group
    form_class = GroupForm
    template_name = 'add-role.html'
    success_message = "%(name)s was created successfully"
    success_url = '/roles/'
    permission_required = ('auth.change_group',)

    def form_valid(self, form):
        self.object = form.save()
        for perm in permissions_list:
            if perm in self.request.POST:
                self.object.permissions.add(Permission.objects.get(codename=perm))
        return HttpResponseRedirect(self.get_success_url())

    def get_context_data(self, **kwargs):
        context = super(RoleCreateView, self).get_context_data(**kwargs)
        context['activate'] = 'users'
        context['is_arc'] = is_arc(self.request.session.get('country'))
        return context


class RoleDetailView(SuccessMessageMixin, PermissionRequiredMixin, UpdateView):
    """ Role edit view."""

    form_class = GroupForm
    model = Group
    template_name = 'role-details.html'

    context_object_name = 'object'
    success_url = '/roles/'
    success_message = "%(name)s was updated successfully"
    permission_required = ('auth.view_group',)

    def form_valid(self, form):
        self.object = form.save()
        self.object.permissions.clear()
        for perm in permissions_list:
            if perm in self.request.POST:
                self.object.permissions.add(Permission.objects.get(codename=perm))
        return HttpResponseRedirect(self.get_success_url())

    def test_func(self):
        return True

    def get_context_data(self, **kwargs):
        context = super(RoleDetailView, self).get_context_data(**kwargs)
        context['activate'] = 'users'
        context['grp_perms'] = self.get_form().instance.permissions.all().values_list('codename', flat=True)
        return context


class RoleUpdateView(SuccessMessageMixin, PermissionRequiredMixin, UpdateView):
    """ Role edit view."""

    form_class = GroupForm
    model = Group
    template_name = 'update-role.html'

    context_object_name = 'object'
    success_url = '/roles/'
    success_message = "%(name)s was updated successfully"
    permission_required = ('auth.change_group',)

    def form_valid(self, form):
        self.object = form.save()
        self.object.permissions.clear()
        for perm in permissions_list:
            if perm in self.request.POST:
                self.object.permissions.add(Permission.objects.get(codename=perm))
        return HttpResponseRedirect(self.get_success_url())

    def test_func(self):
        return True

    def get_context_data(self, **kwargs):
        context = super(RoleUpdateView, self).get_context_data(**kwargs)
        context['activate'] = 'users'
        context['grp_perms'] = self.get_form().instance.permissions.all().values_list('codename', flat=True)
        context['is_arc'] = is_arc(self.request.session.get('country'))
        return context


class RoleDeleteView(PermissionRequiredMixin, DeleteView):
    """Role deletion view."""

    model = Group

    success_url = '/roles'
    permission_required = ('auth.change_group',)

    def delete(self, *args, **kwargs):
        self.get_object().delete()
        return HttpResponseRedirect(self.success_url)

    def test_func(self):
        return True

    def get_success_url(self):
        return reverse_lazy('roles')

    def delete(self, request, *args, **kwargs):
        messages.add_message(self.request, messages.SUCCESS,
                             'User role deleted successfully.')
        return super(RoleDeleteView, self).delete(request, *args, **kwargs)


def get_role_name_status(request):
    """ check if group name already exists. """
    if request.method == 'POST' and request.is_ajax():
        mimetype = 'application/json'
        name = request.POST.get('name')
        id = request.POST.get('id')
        if name:
            if id:
                if Group.objects.filter(name__iexact=name).exclude(id=id).exists():
                    data = 'true'
                else:
                    data = 'false'
            else:
                if Group.objects.filter(name__iexact=name).exists():
                    data = 'true'
                else:
                    data = 'false'
    else:
        data = 'fail'
    return HttpResponse(data, mimetype)
