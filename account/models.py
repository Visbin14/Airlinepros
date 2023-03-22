import datetime
from django.contrib.auth.models import AbstractUser, BaseUserManager
from django.db import models
from django.utils.translation import ugettext_lazy as _
from django.urls import reverse
from main.models import Airline, Country, State


class UserManager(BaseUserManager):
    """
    Model manager for User model.
    """

    use_in_migrations = True

    def _create_user(self, email, password, **extra_fields):
        """Create and save a User with the given email and password."""
        if not email:
            raise ValueError('The given email must be set')
        email = self.normalize_email(email)
        user = self.model(email=email, **extra_fields)
        user.set_password(password)
        user.save(using=self._db)
        return user

    def create_user(self, email, password=None, **extra_fields):
        """Create and save a regular User with the given email and password."""
        extra_fields.setdefault('is_staff', False)
        extra_fields.setdefault('is_superuser', False)
        return self._create_user(email, password, **extra_fields)

    def create_superuser(self, email, password, **extra_fields):
        """Create and save a SuperUser with the given email and password."""
        extra_fields.setdefault('is_staff', True)
        extra_fields.setdefault('is_superuser', True)

        if extra_fields.get('is_staff') is not True:
            raise ValueError('Superuser must have is_staff=True.')
        if extra_fields.get('is_superuser') is not True:
            raise ValueError('Superuser must have is_superuser=True.')

        return self._create_user(email, password, **extra_fields)


class User(AbstractUser):
    """
    User model.
    """
    first_name = models.CharField(_('first name'), max_length=30)
    last_name = models.CharField(_('last name'), max_length=150)
    email = models.EmailField(_('email address'), unique=True)
    airline_user = models.BooleanField(default=False)
    sales_user = models.BooleanField(default=False)
    finance_user = models.BooleanField(default=False)
    admin_user = models.BooleanField(default=False)
    airline = models.ForeignKey(Airline, related_name='users', on_delete=models.CASCADE,  null=True, blank=True)
    avatar = models.ImageField(upload_to='avatars/', blank=True, null=True)
    countries = models.ManyToManyField(
        Country, blank=True, related_name='users')
    own_states = models.ManyToManyField(
        State,
        verbose_name=_('owned states'),
        blank=True,
        help_text=_('Owned states for this user.'),
        related_name="user_set",
        related_query_name="user",
    )
    username = None
    USERNAME_FIELD = 'email'
    REQUIRED_FIELDS = []

    objects = UserManager()

    def __str__(self):
        return self.email

    def get_absolute_url(self):
        return reverse('user_details', kwargs={'pk': self.pk})

    @property
    def photo_url(self):
        """ returns the user avatar url or a place holder avatar. """
        if self.avatar and hasattr(self.avatar, 'url'):
            return self.avatar.url
        else:
            return "/static/main/img/user_placeholder.jpg"