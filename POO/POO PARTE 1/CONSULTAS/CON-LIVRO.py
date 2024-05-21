import openpyxl

# CARREGAR ARQUIVO
trabalho = openpyxl.load_workbook('POO.xlsx')

# SELECIONAR PÁGINA LIVRO
livro_page = trabalho['LIVROS']

# MOSTRAR DADOS DE CADA LINHA
for row in livro_page.iter_rows(min_row=3, min_col=1, max_col=5):
    if row[4].value is None:            # Verifica se a célula na coluna E (index 4) está vazia
        break                           # Interrompe o loop se encontrar uma célula vazia na coluna E
    for cell in row:
        print(cell.value, end=", ")     # Imprime o valor da célula seguido por uma vírgula
    print()                              # Adiciona uma nova linha após cada linha de dados