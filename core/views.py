
from django.shortcuts import render
from django.http import HttpResponse
import os
from analisemercado import gerar_relatorio

def home(request):
    return render(request, 'core/home.html')

def analise_mercado(request):
    if request.method == "POST":
        setor = request.POST.get("categoria", "").strip()
        empresa = request.POST.get("setor", "").strip() or "Todos os setores"
        regiao = request.POST.get("regiao", "").strip() or "Brasil"

        try:
            caminho_arquivo = gerar_relatorio(setor, regiao, empresa)
            with open(caminho_arquivo, 'rb') as f:
                response = HttpResponse(f.read(), content_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
                response['Content-Disposition'] = f'attachment; filename="{os.path.basename(caminho_arquivo)}"'
                return response
        except Exception as e:
            return render(request, 'core/analise.html', {'erro': f"Erro ao gerar relatório: {e}"})

    return render(request, 'core/analise.html')
