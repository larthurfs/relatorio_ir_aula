from django.core.files import File
from django.http import HttpResponse
from django.shortcuts import render
import io
from django.contrib import messages
from django.views.decorators.csrf import csrf_exempt

from Apps.core.excel import criar_planilha_modelo, criar_planilha
from Apps.core.relatorioir import relatorioir


# Create your views here.
def home(request):
    """
    Retorna a pagina inicial do projeto
    """
    return render(request, 'index.html')


def modelo_excel(request):
    if request.method == "GET":
        output = io.BytesIO()

        criar_planilha_modelo(output)

        output.seek(0)

        filename = 'modelo_relatorio.xlsx'

        response = HttpResponse(output, content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

        response['Content-Disposition'] = f'attachment; filename={filename}'

        return response

    return render(request, 'index.html')

@csrf_exempt
def criarrelatorio(request):
    """
    Responsável por chamas as funções de ETL para a criação do relatório de IR e desenvolver um Excel com os dados tratados
    """
    ano = request.POST.get("ano")
    if request.method == "POST":
        dados = File(request.FILES['excel_file'])
        if not dados.name.endswith('xlsx'):
            msn = "Formato de arquivo invalido! Só é aceito arquivos .xlsx"
            messages.error(request, msn)
            return render(request, 'index.html')

        output = io.BytesIO()
        try:
            df = relatorioir(int(ano), dados)
            criar_planilha(df, dados, output)
        except:
            msn = "Verifique se a sua planilha está seguindo o modelo fornecido"
            messages.error(request, msn)
            return render(request, 'index.html')

        output.seek(0)

        filename = 'relatorio_ir.xlsx'

        response = HttpResponse(output,
                                content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

        response['Content-Disposition'] = f'attachment; filename={filename}'

        return response
    return render(request, 'index.html')