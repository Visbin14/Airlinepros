from django.contrib.auth.mixins import UserPassesTestMixin
from django.contrib.auth.views import redirect_to_login


class PermissionCheckMixin(UserPassesTestMixin):
    raise_exception = True

    def dispatch(self, request, *args, **kwargs):
        if not self.request.user.is_authenticated:
            return redirect_to_login(self.request.get_full_path(), self.get_login_url(), self.get_redirect_field_name())
        return super().dispatch(request, *args, **kwargs)

    def test_func(arg):
        return True
