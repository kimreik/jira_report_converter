
from django.http import HttpResponse
from django.template import loader
from django.views.decorators.csrf import csrf_exempt

from .ExcelConverter import ExcelConverter


@csrf_exempt
def index(request):
    if request.method == 'POST':
        file = request.FILES['file']

        response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

        converter = ExcelConverter(file)
        converter.convert(response)

        # Create the HttpResponse object with the appropriate CSV header.
        response['Content-Disposition'] = 'attachment; filename= "report.xlsx"'

        return response

    template = loader.get_template('jira_report_converter/index.html')
    context = {
        'name': request.GET.get('name', 'guest'),
    }
    return HttpResponse(template.render(context, request))
