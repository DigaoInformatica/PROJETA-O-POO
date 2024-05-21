import openpyxl

# Carregar o arquivo existente
trabalho = openpyxl.load_workbook('POO.xlsx')

# Selecionar a aba 'USUARIOS'
usuarios_page = trabalho['USUARIOS']

# Função para adicionar dados com input
def adicionar_usuario():
    dados = [
        input("Digite o nome do usuário: "),
        input("Digite o e-mail do usuário: "),
        input("Digite o cargo do usuário: "),
        input("Digite o departamento do usuário: "),
        input("Digite o status do usuário (Ativo/Inativo): ")
    ]

    linha = 3
    while usuarios_page.cell(row=linha, column=1).value is not None:
        linha += 1

    for col, valor in enumerate(dados, start=1):
        usuarios_page.cell(row=linha, column=col, value=valor)

# Loop para adicionar vários usuários
while True:
    adicionar_usuario()
    if input("Deseja adicionar outro usuário? (s/n): ").strip().lower() != 's':
        break

# Salvar o arquivo
trabalho.save('POO.xlsx')
print("Dados salvos com sucesso.")
