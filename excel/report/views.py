from django.http import FileResponse, HttpResponse
from django.shortcuts import render

from report.forms import ReportForm
from report.servise import ReportService


def correct_report(request, *args, **kwargs):
    if request.method == 'POST':
        file = request.FILES['file']
        if file.name.endswith('.xlsx'):
            file_path = ReportService(file).create_report()
            return FileResponse(open(file_path, 'rb'))

        return HttpResponse(status=400, content='Only .xlsx files are')

    else:
        form = ReportForm()
        return render(request, 'report.html', {'form': form})
