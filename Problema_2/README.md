# Case Python

Problema 2

## Descrição da aplicação 

Diante do desafio de realizar o tratamento de dados em conjunto com a criação de uma interface para inserção 
de dados, desenvolvi a seguinte solução:

O script "Dados_Tratamento.py" trata os dados da planilha original e, opcionalmente, remove a senha do arquivo 
de entrada. Após carregar o arquivo, percorre as duas abas do Excel, realizando um tratamento específico em cada 
coluna para adequar os dados ao formato necessário para o framework Django carregar. 

O Django foi escolhido como a interface principal devido à sua relevância no mercado e facilidade de criação de 
projetos estruturados. Com isso, foi desenvolvido um fluxo completo de URLs, rotas, modelos e templates HTML para que 
o usuário possa cadastrar um novo contaminante, além de acessar uma segunda página com os contaminantes cadastrados 
no banco de dados.

Com a base de dados estabelecida para o relatório final, o arquivo "Dados_saida.py" realiza o último procedimento 
de formatação do Excel com base na imagem fornecida. Em seguida, são aplicadas as condições solicitadas nos itens 
2.1, 2.2, 2.3 e 2.4 para os campos de análise de resultados.

## Requisitos (Bibliotecas Externas)

- Django
- Pandas
- openpyxl
- Django import export

## Como Instalar

- Certifique-se de ter instalado as bibliotecas listadas acima.
- Copie a pasta projeto_ex_2, nela foi criada o projeto des desafio e seu respctivo app. 
- Copie os arquivos .py necessários para a execução do projeto. Estes arquivos são Dados_saida.py e Dados_Tratamento.py.
- Crie as seguintes pastas no diretório onde os arquivos .py foram copiados:
    Dados_Ext
    Dados_Input

## Como Utilizar

Apos realizar os processos de "Como Instalar", coloque os dados na pasta "Dados_Input" e execute o arquivo Dados_Tratamento.py,
em seguência dentro da pasta projeto_ex_2 (no terminar) pode-se executar o comando "python .\manage.py runserver" clique no IP
e entrara na pagina "HOME" do sistema na qual pode fazer o cadastro de um novo elemento, e se clicar no link "Lista de Cadastro" 
será redirecionado para a listagem de todos os comtaminates cadastrados. 

Em seguida execute o arquivo Dados_saida.py, assim o excel final será gerado na pasta raiz com o nome "relatorio.xlsx".
