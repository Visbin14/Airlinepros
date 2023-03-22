from django.db import models

from agency.models import Agency
from report.models.common import ReportFile, ReportPeriod
from postgres_copy import CopyManager
from main.models import Airline
from report.utils import get_float
from report.regex import  disp_record, arc_tot, arc_deduc, arc_fees, arc_rev, arc_net,disp_record1


class Transaction(models.Model):
    agency = models.ForeignKey(Agency, on_delete=models.CASCADE)
    report = models.ForeignKey(ReportFile, on_delete=models.CASCADE)

    transaction_type = models.CharField(max_length=9, null=True, blank=True)
    ticket_no = models.CharField(max_length=99, null=True, blank=True)
    issue_date = models.DateField(null=True, blank=True)

    transaction_amount = models.FloatField(null=True, blank=True, default=0.0)
    fare_amount = models.FloatField(null=True, default=0.0)
    pen = models.FloatField(null=True, default=0.0)
    pen_type = models.CharField(max_length=99, null=True, blank=True)

    cobl_amount = models.FloatField(null=True, default=0.0)
    std_comm_rate = models.FloatField(null=True)
    std_comm_amount = models.FloatField(null=True, default=0.0)
    sup_comm_rate = models.FloatField(null=True, default=0.0)
    sup_comm_amount = models.FloatField(null=True, default=0.0)
    tax_on_comm = models.FloatField(null=True, default=0.0)
    balance = models.FloatField(null=True, default=0.0)

    cpui = models.CharField(max_length=5, null=True, blank=True)
    stat = models.CharField(max_length=5, null=True, blank=True)
    fop = models.CharField(max_length=5, null=True, blank=True)

    cc = models.FloatField(null=True, default=0.0)
    ca = models.FloatField(null=True, default=0.0)
    ep = models.FloatField(null=True, default=0.0)

    card_type = models.CharField(max_length=99, null=True, blank=True)

    #ARC Specific
    card_code = models.CharField(max_length=9, null=True, blank=True)
    international = models.BooleanField(default=False)

    #IATA Specific for arc it will be true always
    is_sale = models.BooleanField(default=True)

    class Meta:
        indexes = [
            models.Index(fields=['agency', 'report', 'transaction_type']),
            models.Index(fields=['transaction_type','ticket_no','issue_date','transaction_amount','fare_amount','pen','pen_type','cobl_amount','std_comm_rate','std_comm_amount','sup_comm_rate','sup_comm_amount','tax_on_comm','balance','cpui','stat','fop','cc','ca','ep','card_type','card_code','international'])
        ]
        ordering = ['id']

    def __str__(self):
        return "%s - %s" % (self.ticket_no, self.agency.trade_name)


class Charges(models.Model):
    amount = models.FloatField(null=True)
    type = models.CharField(max_length=10, null=True, blank=True)
    transaction = models.ForeignKey(Transaction, on_delete=models.CASCADE, null=True)

    class Meta:
        indexes = [
            models.Index(fields=['transaction', 'amount', 'type']),
        ]

    def __str__(self):
        return "%s - %s" % (self.transaction, self.amount)


class Taxes(models.Model):
    amount = models.FloatField()
    type = models.CharField(max_length=10, null=True, blank=True)
    transaction = models.ForeignKey(Transaction, on_delete=models.CASCADE, null=True)

    class Meta:
        indexes = [
            models.Index(fields=['transaction', 'amount', 'type']),
        ]

    def __str__(self):
        return "%s - %s" % (self.transaction, self.amount)


class AgencyDebitMemo(models.Model):
    amount = models.FloatField(null=True)
    allowed_commission_amount = models.FloatField(null=True, default=0.00)
    comment = models.TextField()
    transaction = models.ForeignKey(Transaction, on_delete=models.CASCADE, null=True, unique=True)
    is_acm = models.BooleanField(default=False)

    csv = CopyManager()
    objects = models.Manager()

    class Meta:
        indexes = [
            models.Index(fields=['transaction']),
        ]

    def __str__(self):
        return "%s - %s" % (self.transaction, self.amount)


class Remittance(models.Model):
    ped = models.DateField(unique=True)
    remittance = models.DateField(unique=True)

    class Meta:
        indexes = [
            models.Index(fields=['ped']),
        ]

    def __str__(self):
        return "%s - %s" % (self.ped, self.remittance)


class Disbursement(models.Model):
    report_period = models.ForeignKey(
        ReportPeriod, on_delete=models.CASCADE, null=True)
    airline       = models.ForeignKey(Airline, on_delete=models.CASCADE, null=True)

    filedate      = models.DateField()
    rundate1      = models.DateField()
    rundate2      = models.DateField(null=True)

    file1         = models.FileField(upload_to="disbursements")
    file2         = models.FileField(upload_to="disbursements", null=True)

    # bank4         = models.FloatField()
    # bank5         = models.FloatField()
    bank7         = models.FloatField(default=0.00)
    arc_deduction = models.FloatField(default=0.00)
    arc_fees	  = models.FloatField(default=0.00)
    arc_tot       = models.FloatField(default=0.00)
    arc_reversal  = models.FloatField(default=0.00)
    arc_net_disb  = models.FloatField(default=0.00)

    imported_at   = models.DateTimeField(auto_now_add=True)

    pending_deductions = models.BooleanField(default=False)

    class Meta:
        unique_together = (
            ('airline', 'report_period'),
            ('airline', 'filedate'),
        )

    def __str__(self):
        return "%s - %s" % (self.airline, self.report_period)

    def is_filed(self, date, file):
        if self.rundate2:
            return ('2 reports have already been filed.')
        if date.date() == self.rundate1:
            return ('This file has already been imported.')
        self.file2 = file.split('media/')[-1]
        self.rundate2 = date
        self.save()

    def add_charges(self, bk7, arc, fee, tot, rev, net):
        # self.bank4 += bk4
        # self.bank5 += bk5
        self.bank7 += bk7
        self.arc_deduction += arc
        self.arc_fees += fee
        self.arc_tot += tot
        self.arc_reversal += rev
        self.arc_net_disb += net
        self.save()

    # def get_bankval(self, num):
    #     if num == 4:
    #         return self.bank4
    #     elif num == 5:
    #         return self.bank5
    #     elif num == 7:
    #         return self.bank7

    def disb_total(self):
        tot = self.bank7 - (self.arc_deduction+ self.arc_fees + self.arc_tot + self.arc_reversal)
        if tot < 0:
            return 0
        else:
            return tot

    def reprocess_files(self):
        # self.bank4 = 0
        # self.bank5 = 0
        self.bank7 = 0
        self.arc_deduction = 0
        self.arc_fees = 0
        self.arc_tot = 0
        self.arc_reversal = 0
        self.arc_net_disb = 0
        if self.file1:
            self.process_file(self.file1)
        if self.file2:
            self.process_file(self.file2)

    def process_file(self, dis_file):
        bankpay = 0.00
        pending_payments = False
        tot = '0.00'
        deduc = '0.00'
        fees = '0.00'
        rev = '0.00'
        net = '0.00'
        for i in dis_file.readlines():
            i = i.decode('utf-8')
            i = i.replace('\r', '')
            m = disp_record.match(i)
            if not m:       
                m = disp_record1.match(i)
            if m:
                try:
                    (index, amount,temp_val) = m.groups()
                except:
                    (index, amount) = m.groups()
                bankpay = bankpay + get_float(amount)
                continue

            m = arc_tot.match(i)
            if m:
                (tot, pending) = m.groups()
                pending_payments |= pending == 'NA'
                if pending: tot = '0.0'
                # logger.debug("Parsed arc_tot '%s' pending='%s'"
                #              % (tot, pending == 'NA'))
                continue

            m = arc_deduc.match(i)
            if m:
                (deduc, pending) = m.groups()
                pending_payments |= pending == 'NA'
                if pending: deduc = '0.0'
                # logger.debug("Parsed arc_deduc '%s' pending='%s'"
                #              % (deduc, pending == 'NA'))
                continue

            m = arc_fees.match(i)
            if m:
                (fees, pending) = m.groups()
                pending_payments |= pending == 'NA'
                if pending: fees = '0.0'
                # logger.debug("Parsed arc_fees '%s' pending='%s'"
                #              % (fees, pending == 'NA'))
                continue

            m = arc_rev.match(i)
            if m:
                (rev, pending) = m.groups()
                pending_payments |= pending == 'NA'
                if pending: rev = '0.0'
                # logger.debug("Parsed arc_rev '%s' pending='%s'"
                #              % (rev, pending == 'NA'))
                continue

            m = arc_net.match(i)
            if m:
                (net, pending) = m.groups()
                pending_payments |= pending == 'NA'
                if pending: net = '0.0'
                # logger.debug("Parsed arc_rev '%s' pending='%s'"
                #              % (rev, pending == 'NA'))
                continue

        self.pending_deductions |= pending_payments
        self.add_charges(bankpay, get_float(deduc), get_float(fees), get_float(tot),
            get_float(rev),get_float(net))
        self.save()


class CarrierDeductions(models.Model):
    report_period = models.ForeignKey(
        ReportPeriod, on_delete=models.CASCADE, null=True)
    airline = models.ForeignKey(Airline, on_delete=models.CASCADE, null=True)

    filedate      = models.DateField()
    file          = models.FileField(upload_to="deductions")

    imported_at   = models.DateTimeField(auto_now_add=True)
    no_bill_items = models.IntegerField(null=True)

    class Meta:
        unique_together = ('airline', 'report_period')

    def __str__(self):
        return "%s - %s" % (self.airline, self.report_period)

class Deduction(models.Model):
    report = models.ForeignKey(CarrierDeductions, on_delete=models.CASCADE, null=True)
    type   = models.CharField(max_length=20)
    amount = models.FloatField(null=True)
    pending = models.BooleanField(default=False)

    def __str__(self):
        return "%s - %s" % (self.report, self.type)