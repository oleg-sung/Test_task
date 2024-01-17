from django.urls import path

from report.views import correct_report

app_name = 'report'

urlpatterns = [
    path('report/', correct_report, name='report'),
]
