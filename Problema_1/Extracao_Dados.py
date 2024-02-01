# Importações
import tabula
import os

# Função para extrair dados de acordo com a área especificada e salvar em arquivos CSV
def extrair_e_salvar_dados(area, area_numero, arquivo_input):
    for idx, arquivo in enumerate(arquivo_input):
        arquivo_out = tabula.read_pdf(arquivo, pages="all", area=area)
        for i, df in enumerate(arquivo_out):
            df.to_csv(f"Dados_Ext/arquivo_out_{area_numero}_{idx}_{i}.csv", index=False)

# Variáveis
pasta = "Dados_Input"
arquivo_input = []

# Percorrer a pasta 
for arquivo in os.listdir(pasta):
    if arquivo.endswith('.pdf'):
        caminho_arquivo = os.path.join(pasta, arquivo)
        arquivo_input.append(caminho_arquivo) 
    else:
        pass       

# Definição das áreas de extração
areas = {
    "Area_1": [20, 5, 70, 3000],
    "Area_2": [80, 5, 115, 2500],
    "Area_3": [120, 5, 170, 2500],
    "Area_4": [175, 5, 5000, 2500]
}

for area_nome, area_valor in areas.items():
    extrair_e_salvar_dados(area_valor, area_nome, arquivo_input)



