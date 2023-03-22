from django.db import models




class DashboardModel(models.Model):
    Region = models.CharField(max_length=100,blank=True,null=True)
    Country = models.CharField(max_length=100,blank=True,null=True)
    Year = models.CharField(max_length=100,blank=True,null=True)
    Airline = models.CharField(max_length=100,blank=True,null=True)
    Airline_IATA_Code = models.CharField(max_length=50,blank=True,null=True)
    IATA_Numeric_Code = models.IntegerField()
    Agreement_Type = models.CharField(max_length=50,blank=True,null=True)
    Contraced_With = models.CharField(max_length=500,blank=True,null=True)
    Month = models.CharField(max_length=100,blank=True,null=True)
    ORC_Rate = models.FloatField(blank=True,null=True)
    Gross = models.FloatField(blank=True,null=True)
    Nett = models.FloatField(blank=True,null=True)
    ORC_Actual = models.FloatField(blank=True,null=True)



    def __str__(self):
        return self.Country +" "+ self.Year
