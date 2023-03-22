from django import forms
from report.models import ReportFile


class ReportFileForm(forms.ModelForm):

    class Meta:
        model = ReportFile
        fields = ('file',)
