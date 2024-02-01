# Case Python

Problema 1.

## Descrição da aplicação 

Dado a problematica de realizar a leitura de dados não editaveis (pdf), foi elaborado2 aquivos .py (Extracao_Dados.py e Analise_Dados.py):

Extracao_Dados.py = Ao passar o caminha da pasta "Dados_Input" independente do quantidade e 
    nomenclatura dos arquivo o script irá conseguir extrair os dados contidos. No modelo PDF passado a extração das informações são passadas para 4 database, na busca de uma extração mais eficiente
    e precisa. 
    
Exemplo a area contendo os parametro de 1 PDF sera salva como  arquivo_out_Area_4_0_0.csv 
    sendo 4 a area respectiva, primeiro 0 para a seguencia do arquivo e 3 o index, isso na pasta Dados_Ext.

Analise_Dados = O script percorre a pasta Dados_Ext garregando todas os arquivos com suas areas relacionadas,
    apos percorrer realiza um tratamento nos dados necessario para cada tabela e posteriormente é realizado a criação do 
    excel final e sua devida formatação. 

A pasta "Dados_Analise" foi criada com intuito de uma melhor visualização do processo para a geração do Dataframe_Final.

## Requisitos (Bibliotecas Externas)

- Tabula
- Pandas
- openpyxl

## Como Instalar

- Certifique-se de ter instalado as bibliotecas listadas acima 
- Copie os arquivos .py necessários para a execução do projeto. Estes arquivos são Extracao_Dados.py e Analise_Dados.py.
- Crie as seguintes pastas no diretório onde os arquivos .py foram copiados:
    Dados_Analise
    Dados_Ext
    Dados_Input

## Como Utilizar

Apos realizar os processos de "Como Instalar", coloque os dados na pasta "Dados_Input" e execute o arquivo Extracao_Dados.py,
em seguência pode-se executar o arquivo Analise_Dados.py, o excel final será gerado na pasta raiz com o nome "Dataframe_Final.xlsx".

* Caso aparecer um erro "Please ensure Java", será necessário instalar o Java no computador, de acordo com o link:
"https://stackoverflow.com/questions/54817211/java-command-is-not-found-from-this-python-process-please-ensure-java-is-inst"
