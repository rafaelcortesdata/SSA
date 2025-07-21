import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime

SERVICE_ACCOUNT_FILE =  r"C:\Users\Administrador\Documents\meus scripts\projeto-dados-suporte-copiai.json"   # Caminho do seu JSON
SPREADSHEET_NAME = "RELATORIO_SUPORTE_COPIEAI"          # Nome da sua planilha

scope = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

credentials = ServiceAccountCredentials.from_json_keyfile_name(
    SERVICE_ACCOUNT_FILE, scope
)
client = gspread.authorize(credentials)
sheet = client.open(SPREADSHEET_NAME).sheet1

COLUNAS = [
    "Cancelados", "Reembolso", "Login", "Cadastro",
    "Clonagem", "Espião", "Dúvidas Gerais", "Bugs",
    "Domínio", "Pixel", "Respondidos"
]

def registrar_clique(nome_botao):
    """
    Atualiza a contagem de cliques na célula correta.
    Se a data não existir, cria nova linha.
    """
    hoje = datetime.now().strftime("%d/%m/%Y")

    datas = sheet.col_values(1)

    if hoje in datas:
        row_index = datas.index(hoje) + 1
    else:
        nova_linha = [hoje] + ["0"] * len(COLUNAS)
        sheet.append_row(nova_linha)
        row_index = len(datas) + 1

    try:
        col_index = COLUNAS.index(nome_botao) + 2
    except ValueError:
        print(f"Botão '{nome_botao}' não encontrado no header.")
        return

    valor_atual = sheet.cell(row_index, col_index).value
    if valor_atual is None or valor_atual == "":
        valor_atual = "0"

    novo_valor = str(int(valor_atual) + 1)

    sheet.update_cell(row_index, col_index, novo_valor)
