from django.shortcuts import render
from .models import Cadastro

def home(request):
    return render(request, 'usuarios/home.html')

def cadastro(request):
    if request.method == 'POST':
        novo_cadastro = Cadastro.objects.create(
            contaminante = request.POST.get('contaminante'),
            cas = request.POST.get('cas'),
            vor = request.POST.get('vor'),
            vor_valor = request.POST.get('vor_valor'),
            abe_c = request.POST.get('abe_c'),
            abe_nc = request.POST.get('abe_nc'),
            fec_c = request.POST.get('fec_c'),
            fec_nc = request.POST.get('fec_nc')
        )

        # Retornar os dados para a pagina de listagem de cadastro
        cadastros = Cadastro.objects.all()
        return render(request, 'usuarios/cadastrados.html', {'cadastros': cadastros})
    
    else:
            cadastros = Cadastro.objects.all()
            return render(request, 'usuarios/cadastrados.html', {'cadastros': cadastros})
