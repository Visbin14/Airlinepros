from django.db import models
from django.conf import settings
from postgres_copy import CopyManager
from django.urls import reverse


from main.models import BaseModel, Country, State, City,Airline

# Create your models here.


class IncompleteAgencyError(Exception):
    pass


# modes for angency to be in
STATUS_MODES = (
    ('A', 'Active'),
    ('D', 'Defaulted'),
    ('R', 'Revoked'),
    ('S', 'Reinstated'),
    ('T', 'Terminated'),
)

STATUS_MODES_IATA = (
    ('A', 'Active'),
    ('D', 'Default Information'),
    ('R', 'Reviews/Notices of Termination'),
    ('S', 'Reinstatements'),
    ('T', 'Terminations And Closures'),
    ('I', 'Irregularities And Admin Noncompliance'),
)

# modes for angency to be in
FILE_TYPE_CHOICES = (
    (1, 'Agency List'),
    (2, 'Revokation/Reinstatment'),
)


class RRTable(BaseModel):
    """
    holds the revokation/reinstatment bullitins as the date issued
    """
    date = models.DateField()
    file = models.FileField(upload_to='uploadedfiles/')


class AgencyType(BaseModel):
    """
    Model to monitor the angency types.
    """
    country = models.ForeignKey(
        Country, null=True, blank=True, related_name='agencytype', on_delete=models.SET_NULL)
    name = models.CharField(max_length=300, null=True)

    class Meta:
        unique_together = ('name', 'country')

    def __str__(self):
        return self.name

    def get_absolute_url(self):
        return reverse('agency_type_details', kwargs={'pk': self.pk})


class AgencyCollection(BaseModel):
    """
    Model to monitor the angency collections.
    """
    country = models.ForeignKey(
        Country, null=True, blank=True, related_name='agencycollections', on_delete=models.SET_NULL)
    name = models.CharField(max_length=300, null=True)

    class Meta:
        unique_together = ('name', 'country')

    def __str__(self):
        return self.name

    def get_absolute_url(self):
        return reverse('agency_collection_details', kwargs={'pk': self.pk})


class Agency(BaseModel):
    """
    Agency model, holds all the relevant agency details, got from agency
    list file. Only required field is the agency number.
    """
    agency_no = models.CharField(max_length=99, null=True, blank=True)
    trade_name = models.CharField(max_length=199, null=True)
    address1 = models.CharField(max_length=99, null=True, blank=True)
    address2 = models.CharField(max_length=99, null=True, blank=True)
    city = models.ForeignKey(
        City, on_delete=models.SET_NULL, null=True, blank=True)
    state = models.ForeignKey(
        State, on_delete=models.SET_NULL, null=True, blank=True)
    country = models.ForeignKey(
        Country, null=True, blank=True, related_name='agencies', on_delete=models.SET_NULL)
    zip_code = models.CharField(max_length=9, null=True, blank=True)
    email = models.EmailField(null=True, blank=True)
    vat_number = models.CharField(max_length=50, null=True, blank=True)
    tel = models.CharField(max_length=16, null=True, blank=True)
    agency_type = models.ForeignKey(
        AgencyType, null=True, blank=True, related_name='agencies', on_delete=models.SET_NULL)
    agency_collection = models.ForeignKey(
        AgencyCollection, null=True, blank=True, related_name='collection_agencies', on_delete=models.SET_NULL)
    home_agency = models.CharField(max_length=99, null=True, blank=True)
    sales_owner = models.ForeignKey(settings.AUTH_USER_MODEL, limit_choices_to={
                                    'is_active': True}, null=True, blank=True, on_delete=models.SET_NULL, related_name='owned_agencies')
    status = models.CharField(max_length=1, null=True,
                              blank=True, choices=STATUS_MODES, default='A')
    status_iata = models.CharField(max_length=1, null=True,
                                   blank=True, choices=STATUS_MODES_IATA, default='A')
    from_old_asplinks = models.BooleanField(default=False)  # temporary field to identify the agencies added from old
                                                            # asplinks database

    csv = CopyManager()
    objects = models.Manager()

    class Meta:
        unique_together = ('agency_no', 'country')
      
        permissions = (
            ('download_agencylist', "Can download agency list"),
        )

    def __str__(self):
        return str("%s-%s-%s-%s" % (self.trade_name, self.city, self.agency_no,self.country.name if self.country else ""))

    def get_absolute_url(self):
        return reverse('agency_details', kwargs={'pk': self.pk})

    def update_data(self, line):
        """
        Called to update the details of an agency from the agency list.
        Details are extracted from a line passed from the list.
        parameters are determined by the list spec file avialable
        from the ARC website.
        """
        if len(line) < 694:
            raise IncompleteAgencyError()

        self.trade_name = line[9:34].strip()
        self.address1 = line[87:117].strip()
        self.address2 = line[117:147].strip()
        self.city = line[147:160].strip()
        self.state = line[160:162].strip()
        self.zip_code = line[162:171].strip()
        self.tel = line[171:181].strip()
        self.vat_number = line[198:207].strip()
        self.email = line[368:448].strip()

    # def import_data(self, line, dt, csv=False):
    #     """
    #     Called for a new agency when importing the angency List.
    #     Sets the status    of the agency as active.
    #     """
    #     if csv:
    #         self.agency_no = extract_number(line[1])
    #         results = self.update_csv(line)
    #     else:
    #         self.agency_no = int(line[1:9])
    #         result = self.update_data(line)
    #
    #     x = StatusHistory(date=dt)
    #     x.save()
    #     x.ChangestatustoActive(dt, 'From Agency List')
    #     self.status_history = x
    #
    # def update_csv(self, data):
    #     self.trade_name = data[2].strip()
    #     self.address1 = data[5].strip()
    #     self.address2 = data[6].strip()
    #     self.city = data[7].strip()
    #     self.state = data[8].strip()
    #     self.zip_code = extract_number(data[9])
    #     self.tel = data[10].strip()
    #     self.vat_number = data[14].strip()
    #     self.email = data[33].strip()


class StatusChange(BaseModel):
    """
    Model to monitor changes in the angency status.
    """
    old_status = models.CharField(max_length=99)
    new_status = models.CharField(max_length=99)
    reason = models.CharField(max_length=300, null=True)
    agency = models.ForeignKey(
        Agency, related_name='status_changes', on_delete=models.CASCADE)

    def __str__(self):
        return "%s %s - %s" % (self.agency,self.old_status, self.new_status)


class AgencyContacts(BaseModel):
    """
    Model to hold contacts for each agency.
    """
    name = models.CharField(max_length=25, null=True)
    email = models.EmailField()
    agency = models.ForeignKey(Agency, on_delete=models.CASCADE)


class ListOfAgents(models.Model):
    title = models.CharField(max_length=100, unique=True)
    agents = models.ManyToManyField(Agency)

    def __str__(self):
        return self.title

    class Meta:
        ordering = ('title',)


class AgencyListReference(models.Model):
    """
    model to hold the angency list file and a reference to ensure that the
    the list is not imported again
    """
    processed_at = models.DateTimeField(unique=True, auto_now=True,)
    created_at = models.DateTimeField(auto_now_add=True)
    file = models.FileField(upload_to='agencyfiles/')
    file_type = models.IntegerField(choices=FILE_TYPE_CHOICES, default=1)




class Agency2(BaseModel):
    """
    Agency model, holds all the relevant agency details, got from agency
    list file. Only required field is the agency number.
    """
    agency_no = models.CharField(max_length=99, null=True, blank=True)
    trade_name = models.CharField(max_length=199, null=True)
    address1 = models.CharField(max_length=99, null=True, blank=True)
    address2 = models.CharField(max_length=99, null=True, blank=True)
    city = models.ForeignKey(
        City, on_delete=models.SET_NULL, null=True, blank=True)
    state = models.ForeignKey(
        State, on_delete=models.SET_NULL, null=True, blank=True)
    country = models.ForeignKey(
        Country, null=True, blank=True, related_name='agencies2', on_delete=models.SET_NULL)
    zip_code = models.CharField(max_length=9, null=True, blank=True)
    email = models.EmailField(null=True, blank=True)
    vat_number = models.CharField(max_length=50, null=True, blank=True)
    tel = models.CharField(max_length=16, null=True, blank=True)
    agency_type = models.ForeignKey(
        AgencyType, null=True, blank=True, related_name='agencies2', on_delete=models.SET_NULL)
    agency_collection = models.ForeignKey(
        AgencyCollection, null=True, blank=True, related_name='collection_agencies2', on_delete=models.SET_NULL)
    home_agency = models.CharField(max_length=99, null=True, blank=True)
    sales_owner = models.ForeignKey(settings.AUTH_USER_MODEL, limit_choices_to={
                                    'is_active': True}, null=True, blank=True, on_delete=models.SET_NULL, related_name='owned_agencies2')
    status = models.CharField(max_length=1, null=True,
                              blank=True, choices=STATUS_MODES, default='A')
    status_iata = models.CharField(max_length=1, null=True,
                                   blank=True, choices=STATUS_MODES_IATA, default='A'),
    exclude_list_adm = models.BooleanField(default=False),
    from_old_asplinks = models.BooleanField(default=False)  # temporary field to identify the agencies added from old
                                                            # asplinks database
    csv = CopyManager()
    objects = models.Manager()

    class Meta:
        unique_together = ('agency_no', 'country')
        ordering = ['-trade_name']
        permissions = (
            ('download_agencylist', "Can download agency list"),
        )

    def __str__(self):
        return str("%s-%s-%s" % (self.trade_name, self.city, self.agency_no))

    def get_absolute_url(self):
        return reverse('agency_details', kwargs={'pk': self.pk})

    def update_data(self, line):
        """
        Called to update the details of an agency from the agency list.
        Details are extracted from a line passed from the list.
        parameters are determined by the list spec file avialable
        from the ARC website.
        """
        if len(line) < 694:
            raise IncompleteAgencyError()

        self.trade_name = line[9:34].strip()
        self.address1 = line[87:117].strip()
        self.address2 = line[117:147].strip()
        self.city = line[147:160].strip()
        self.state = line[160:162].strip()
        self.zip_code = line[162:171].strip()
        self.tel = line[171:181].strip()
        self.vat_number = line[198:207].strip()
        self.email = line[368:448].strip()


class NewAgents(BaseModel):
    agency = models.ForeignKey(Agency,related_name='new_agency',on_delete=models.CASCADE)
    airline = models.ForeignKey(Airline,related_name='new_agency_airlines',null=True,on_delete=models.SET_NULL)
    ped = models.DateField(null=True,blank=True)
    week = models.CharField(max_length=5,null=True,blank=True)