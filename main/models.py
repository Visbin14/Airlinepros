import uuid

from django.conf import settings
from django.db import models
from django.shortcuts import reverse


COMMISSION_TYPES = (
    ('M', 'Max Sales commission'),
    ('G', 'GSA Commission'),
    ('I', 'IATA Coordination Fee'),
    ('D', 'Distribution Intermediary Fee'),
    ('A', 'ARC coordination fee'),
)


class BaseModel(models.Model):
    """
    model that will be inherited by all models to add common fields.
    """

    created_on = models.DateTimeField(auto_now_add=True, blank=True)
    modified_on = models.DateTimeField(auto_now=True, blank=True)
    created_by = models.ForeignKey(
        settings.AUTH_USER_MODEL, related_name='%(class)s_createdby', on_delete=models.CASCADE, null=True, blank=True)
    modified_by = models.ForeignKey(settings.AUTH_USER_MODEL,
                                    related_name='%(class)s_modifiedby', on_delete=models.CASCADE, null=True,
                                    blank=True)

    class Meta:
        abstract = True


class Country(BaseModel):
    """Country Model to store country details. """

    name = models.CharField(max_length=50)
    code = models.CharField(max_length=2)
    currency = models.CharField(max_length=4)
    flag = models.ImageField(upload_to='flags/', blank=True, null=True)

    @property
    def photo_url(self):
        """ returns the country flag url or a place holder image. """
        if self.flag and hasattr(self.flag, 'url'):
            return self.flag.url
        else:
            return "/static/main/img/flag_placeholder.jpg"

    def __str__(self):
        return self.name


class Airline(BaseModel):
    """Airline Model. """

    code = models.CharField(max_length=3)
    abrev = models.CharField(max_length=2)
    name = models.CharField(max_length=500)
    arc_coordination_fee = models.FloatField(
        "ARC Coordination Fee (%)", default=0.0)
    accepts_AMEX = models.NullBooleanField()
    accepts_MC = models.NullBooleanField()
    accepts_VI = models.NullBooleanField()
    accepts_UATP = models.NullBooleanField()
    product_manager = models.ForeignKey(settings.AUTH_USER_MODEL, related_name='managed_airlines', null=True, blank=True,
                                        on_delete=models.SET_NULL)
    country = models.ForeignKey(
        Country, on_delete=models.CASCADE, null=True, blank=True, related_name="airlines")

    class Meta:
        ordering = ('name',)
        unique_together = (("country", "code"), ("country", "abrev"),)

    def __str__(self):
        return self.name

    def get_absolute_url(self):
        return reverse('airline_details', kwargs={'pk': self.pk})

    def test_card_transaction(self, code):
        if code == 'VI' and self.accepts_VI:
            return True
        if code == 'CA' and self.accepts_MC:
            return True
        if code == 'AX' and self.accepts_AMEX:
            return True
        if code == 'EX' or code == 'TK':
            return True
        else:
            return False


class Airline_Contact(BaseModel):
    """
    Model for airline contact details.
    """
    name = models.CharField(max_length=50)
    email = models.EmailField(null=True)
    tel = models.IntegerField(null=True)
    notes = models.TextField(null=True)
    airline = models.ForeignKey(
        Airline, related_name='contacts', null=True, on_delete=models.CASCADE)

    def __str__(self):
        return self.name


class State(BaseModel):
    """
    Model for access to the state data
    """
    abrev = models.CharField(max_length=2)
    name = models.CharField(max_length=32)
    country = models.ForeignKey(
        Country, on_delete=models.CASCADE, null=True, blank=True, related_name="states")
    owner = models.ForeignKey(settings.AUTH_USER_MODEL, limit_choices_to={
                                 'is_active': True}, related_name='state', null=True, blank=True, on_delete=models.SET_NULL)

    class Meta:
        ordering = ('abrev',)
        permissions = (
            ('view_stateowners', "Can view state owners list"),
            ('change_stateowners', "Can change state owners"),
        )

    def __str__(self):
        return self.name


class City(BaseModel):
    """
    Model for access to the city data
    """
    name = models.CharField(max_length=32)
    country = models.ForeignKey(
        Country, on_delete=models.CASCADE, null=True, blank=True)
    state = models.ForeignKey(
        State, on_delete=models.CASCADE, null=True, blank=True)

    class Meta:
        ordering = ('name',)

    def __str__(self):
        return self.name


class CommissionHistory(BaseModel):
    """
    Model for access to the Airline commission data
    """
    airline = models.ForeignKey(Airline, on_delete=models.CASCADE, null=True)
    type = models.CharField(max_length=1, null=True, blank=True, choices=COMMISSION_TYPES)
    from_date = models.DateField(blank=True, null=True)
    to_date = models.DateField(blank=True, null=True)
    rate = models.FloatField(null=True)
    add_yq_yr = models.BooleanField(default=False)

    class Meta:
        ordering = ('created_on',)

    def __str__(self):
        return self.airline.name+"  -  "+self.type
class RemoteServers(BaseModel):
    hostname = models.CharField(max_length=50,null=True,blank=True)
    user = models.CharField(max_length=50,null=True,blank=True)
    password = models.CharField(max_length=50,null=True,blank=True)
    port = models.CharField(max_length=50,null=True,blank=True)
    countrycode = models.CharField(max_length=50,null=True,blank=True)

    def __str__(self):
        return str(self.user)+"@"+str(self.hostname)


class LatestFiles(BaseModel):
    ftp_obj = models.ForeignKey(RemoteServers,related_name='rs_latestfile',null=True,blank=True,on_delete=models.CASCADE)
    latest = models.CharField(max_length=50,null=True,blank=True)


class FTPhistory(BaseModel):
    ftp_obj = models.ForeignKey(RemoteServers, related_name='ftp_dwnld_history', null=True, blank=True,
                                on_delete=models.CASCADE)
    file = models.CharField(max_length=100,null=True,blank=True)
    status = models.BooleanField(default=False)
