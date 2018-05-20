
from django.http import HttpResponse
from django.template import loader
from django.utils.http import urlquote
from django.views.decorators.csrf import csrf_exempt

from .ExcelConverter import ExcelConverter


@csrf_exempt
def index(request):
    if request.method == 'POST':
        file = request.FILES['file']

        converter = ExcelConverter(file)
        output = converter.convert()
        file_name = converter.get_file_name()
        output.seek(0)

        response = HttpResponse(output.read(), content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        output.close()

        response['Content-Transfer-Encoding'] = 'Binary'
        # Create the HttpResponse object with the appropriate CSV header.
        # response['Content-Disposition'] = 'attachment; filename= "report.xlsx"'
        response['Content-Disposition'] = 'attachment; filename*=UTF-8\'\'' + urlquote(file_name)

        return response

    template = loader.get_template('jira_report_converter/index.html')
    context = {
        'name': request.GET.get('name', 'guest'),
    }
    return HttpResponse(template.render(context, request))
