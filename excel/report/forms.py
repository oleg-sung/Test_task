from django import forms


class ReportForm(forms.Form):
    """
    Форма для загрузки excel файла
    """
    file = forms.FileField()



