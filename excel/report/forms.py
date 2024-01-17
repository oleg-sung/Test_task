from django import forms


class ReportForm(forms.Form):
    file = forms.FileField()



