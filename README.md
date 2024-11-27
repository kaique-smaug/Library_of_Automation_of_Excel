Manipulação de Automação do Excel
Este projeto contém uma classe em Python chamada Excelque oferece funcionalidades para automação de operações em planilhas Excel. A classe utiliza bibliotecas como openpyxl, pandase xlwingspara leitura, escrita, exclusão de dados, execução de macros e outras tarefas relacionadas a planilhas.

Requisitos
Certifique-se de instalar as bibliotecas específicas no arquivo requirements.txt. Para isso, execute o comando abaixo no terminal:

bater

Copiar código
pip install -r requirements.txt
Conteúdo dorequirements.txt
texto simples

Copiar código
pandas
openpyxl
xlwings
unidecode
Funcionalidades
1. Inicialização
A classe é inicializada com o caminho da planilha que será manipulada.

Pitão

Copiar código
excel = Excel("caminho/para/sua/planilha.xlsx")
2. Escrever Valores
Método:receive_values
Escrever valores em colunas especificadas de uma planilha.

Pitão

Copiar código
excel.receive_values(
    nameSheet="NomeDaPlanilha",
    columnOne="A", valuesOne=["valor1", "valor2"],
    columnTwo="B", valuesTwo=["valor3", "valor4"]
)
Método:write_many_values
Permite escrever vários valores em colunas específicas.

Pitão

Copiar código
excel.write_many_values(
    nameSheet="NomeDaPlanilha",
    columnOne="A", valuesOne=["valor1", "valor2"],
    columnTwo="B", valuesTwo=["valor3", "valor4"]
)
3. Excluir Dados
Método:delete_data
Remova todos os valores de uma planilha.

Pitão

Copiar código
excel.delete_data(nameSheet="NomeDaPlanilha")
Método:delete_data_v2
Uma versão alternativa para exclusão de dados a partir da linha 2.

Pitão

Copiar código
excel.delete_data_v2(nameSheet="NomeDaPlanilha")
4. Executar Macros
Método:macro
Executa uma macro VBA armazenada na planilha.

Pitão

Copiar código
excel.macro(module="Modulo1", sub="MinhaMacro")
5. Ler Valores
Método:read_values
Lê valores de colunas específicas de uma planilha e retorna um dicionário.

Pitão

Copiar código
valores = excel.read_values(
    nameSheet="NomeDaPlanilha",
    columnOne="A",
    columnTwo="B"
)
print(valores)
6. Inserir Valores com Transformação
Método:insert_values
Insira valores de uma planilha em outra com transformações, como remoção de acentos.

Pitão

Copiar código
excel.insert_values(nameSheetOne="Planilha1", nameSheetTwo="Planilha2")
Estrutura do Projeto
ClasseExcel : Contém todos os métodos para automação de planilhas.
Dependências : As bibliotecas usadas são:
openpyxlpara manipulação de arquivos Excel.
pandaspara leitura e manipulação de dados.
xlwingspara execução de macros.
unidecodepara remoção de acentos em texto.
Exemplo Completo
Abaixo está um exemplo completo de como usar uma classe Excel:

Pitão

Copiar código
from excel_class import Excel

# Inicialização
excel = Excel("caminho/para/sua/planilha.xlsx")

# Ler valores de uma planilha
valores = excel.read_values(nameSheet="Planilha1", columnOne="A", columnTwo="B")
print(valores)

# Escrever valores em outra planilha
excel.write_many_values(
    nameSheet="Planilha2",
    columnOne="A", valuesOne=["valor1", "valor2"],
    columnTwo="B", valuesTwo=["valor3", "valor4"]
)

# Executar uma macro
excel.macro(module="Modulo1", sub="MinhaMacro")
Considerações Finais
Essa aula foi desenvolvida para facilitar tarefas comuns em planilhas Excel, reduzindo a necessidade de interação manual. Caso encontre problemas ou tenha sugestões, fique à vontade para contribuir.

Autor: Kaique Batista Ramos
Versão: 1.1.5
