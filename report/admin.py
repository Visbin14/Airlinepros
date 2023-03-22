from django.contrib import admin
from report.models import ReportFile, ReportPeriod, Transaction, Charges, Taxes, AgencyDebitMemo, Remittance, DailyCreditCardFile, ReprocessFile, Disbursement, CarrierDeductions, Deduction,ExcelReportDownload
from django.contrib.admin.models import LogEntry

# Register your models here.
class TransactionAdmin(admin.ModelAdmin):
    search_fields = ('agency__agency_no', 'ticket_no', 'transaction_type', )
    list_display = ('ticket_no', 'agency', 'report', 'transaction_type','card_type','card_code')
    list_filter = ('report', 'transaction_type')

@admin.register(LogEntry)
class LogEntryAdmin(admin.ModelAdmin):
	list_display = ('action_time', 'user', 'object_repr', 'action_flag',)

# class ReportPeriodAdmin(admin.ModelAdmin):
#     search_fields = ('country__name', )
#     list_display = ('week','year','month','country',)

admin.site.register(ReportFile)
admin.site.register(ReportPeriod)
admin.site.register(Transaction, TransactionAdmin)
admin.site.register(Taxes)
admin.site.register(Charges)
admin.site.register(AgencyDebitMemo)
admin.site.register(Remittance)
admin.site.register(DailyCreditCardFile)
admin.site.register(ReprocessFile)
admin.site.register(Disbursement)
admin.site.register(CarrierDeductions)
admin.site.register(Deduction)
admin.site.register(ExcelReportDownload)

