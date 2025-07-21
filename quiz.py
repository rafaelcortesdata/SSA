import tkinter as tk
from PIL import Image, ImageTk
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime, timedelta
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt
import os
import sys
import json

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

SERVICE_ACCOUNT_FILE = r"C:\Users\Administrador\Documents\app-support-quiz\dados-suporte-quiz.json"

if not os.path.exists(SERVICE_ACCOUNT_FILE):
    with open(SERVICE_ACCOUNT_FILE, "w", encoding="utf-8") as f:
        json.dump({}, f)

SPREADSHEET_NAME = "RELATORIO_SUPORTE_CAKTOQUIZ"

COLUNAS = [
    "Domínios", "Pixel", "Login", "clonador", "favicom", "Checkout",
    "Link publicado", "Dúvidas gerais", "Outros Setores", "Bugs",
    "Bloqueados", "Leads", "Webhook", "Fluxo",
    "Usabilidade", "Dicas produtos", "Solicitação Funcionalidades",
    "Respondidos"
]

scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
credentials = ServiceAccountCredentials.from_json_keyfile_name(
    SERVICE_ACCOUNT_FILE,
    scope
)

client = gspread.authorize(credentials)
spreadsheet = client.open(SPREADSHEET_NAME)
sheet = spreadsheet.sheet1

def garantir_cabecalho_e_data(sheet):
    valores = sheet.get_all_values()
    if not valores or valores[0][0] != "Data":
        cabecalho = ["Data"] + COLUNAS
        sheet.insert_row(cabecalho, 1)
    hoje = datetime.now().strftime("%d/%m/%Y")
    datas = sheet.col_values(1)
    if hoje not in datas:
        nova_linha = [hoje] + ["0"] * len(COLUNAS)
        sheet.append_row(nova_linha)

garantir_cabecalho_e_data(sheet)

dados_cache = {botao: 0 for botao in COLUNAS if botao != "Respondidos"}
dados_cache["Respondidos"] = 0
atualizacao_pendente = False

def registrar_clique(nome_botao, acao="somar"):
    global atualizacao_pendente
    if nome_botao == "Respondidos":
        return
    if acao == "somar":
        dados_cache[nome_botao] += 1
    elif acao == "subtrair" and dados_cache[nome_botao] > 0:
        dados_cache[nome_botao] -= 1
    elif acao == "zerar":
        dados_cache[nome_botao] = 0

    dados_cache["Respondidos"] = sum(v for k, v in dados_cache.items() if k != "Respondidos")
    atualizar_labels()
    atualizacao_pendente = True

def salvar_na_planilha():
    global atualizacao_pendente, after_id
    try:
        if not atualizacao_pendente:
            after_id = root.after(15000, salvar_na_planilha)  # 15s agora
            return
        hoje = datetime.now().strftime("%d/%m/%Y")
        datas = sheet.col_values(1)
        if hoje not in datas:
            nova_linha = [hoje] + ["0"] * len(COLUNAS)
            sheet.append_row(nova_linha)
            datas = sheet.col_values(1)

        row_index = datas.index(hoje) + 1

        valores_para_atualizar = [[str(dados_cache.get(botao, 0)) for botao in COLUNAS]]
        cell_range = f'B{row_index}:{chr(66 + len(COLUNAS) - 1)}{row_index}'
        sheet.update(cell_range, valores_para_atualizar)

        atualizacao_pendente = False
    except Exception as e:
        print(f"Erro ao salvar na planilha: {e}")
    finally:
        after_id = root.after(15000, salvar_na_planilha)  # agendamento continua

def obter_cliques_do_dia():
    hoje = datetime.now().strftime("%d/%m/%Y")
    datas = sheet.col_values(1)
    if hoje in datas:
        row_index = datas.index(hoje) + 1
        valores = sheet.row_values(row_index)[1:]
        d = {}
        for i, botao in enumerate(COLUNAS):
            if i < len(valores) and valores[i].isdigit():
                d[botao] = int(valores[i])
            else:
                d[botao] = 0
        return d
    return {botao: 0 for botao in COLUNAS}

def atualizar_labels():
    for botao, lbl in labels.items():
        lbl.config(text=f"{dados_cache.get(botao, 0)}")
    atualizar_grafico()

def agendar_zerar_dados_diariamente():
    now = datetime.now()
    proximo_zerar = (now + timedelta(days=1)).replace(hour=0, minute=1, second=0, microsecond=0)
    segundos_ate_zerar = (proximo_zerar - now).total_seconds()
    root.after(int(segundos_ate_zerar * 1000), zerar_dados_diarios)

def zerar_dados_diarios():
    global atualizacao_pendente
    hoje = datetime.now().strftime("%d/%m/%Y")
    datas = sheet.col_values(1)
    if hoje not in datas:
        nova_linha = [hoje] + ["0"] * len(COLUNAS)
        sheet.append_row(nova_linha)

    for botao in dados_cache:
        dados_cache[botao] = 0
    dados_cache["Respondidos"] = 0
    atualizar_labels()
    atualizacao_pendente = True
    agendar_zerar_dados_diariamente()

# === TKINTER ===
root = tk.Tk()
root.title("S_S_A - Suporte Quiz")
root.geometry("1000x850")
root.configure(bg="black")
root.attributes("-topmost", True)

def on_closing():
    global after_id
    if after_id is not None:
        root.after_cancel(after_id)
    root.destroy()

root.protocol("WM_DELETE_WINDOW", on_closing)

top_frame = tk.Frame(root, bg="black")
top_frame.pack(pady=(20, 10))

try:
    logo_path = r"C:\Users\Administrador\Documents\app-support-quiz\logo.png"
    img = Image.open(logo_path)
    img = img.resize((180, 180), Image.Resampling.LANCZOS)
    logo_img = ImageTk.PhotoImage(img)
    logo_label = tk.Label(top_frame, image=logo_img, bg="black")
    logo_label.image = logo_img
    logo_label.pack()
except Exception as e:
    print("Erro ao carregar o logo:", e)

main_frame = tk.Frame(root, bg="black")
main_frame.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)

frame_botoes = tk.Frame(main_frame, bg="black")
frame_botoes.pack(side=tk.LEFT, padx=10, anchor="n")

labels = {}

def criar_linha_botao(nome, pos):
    linha = pos // 2
    coluna_base = (pos % 2) * 4  # 4 colunas por botão com seus controles

    btn_bg = 'lightgray' if nome != "Respondidos" else 'lightgreen'
    btn = tk.Button(frame_botoes, text=nome, width=18, height=2, bg=btn_bg, command=lambda: registrar_clique(nome))
    btn.grid(row=linha, column=coluna_base, padx=5, pady=4)

    lbl = tk.Label(frame_botoes, text="0", width=4, font=('Arial', 12, 'bold'), bg="black", fg="white")
    lbl.grid(row=linha, column=coluna_base + 1, padx=8)
    labels[nome] = lbl

    if nome != "Respondidos":
        tk.Button(frame_botoes, text="Zerar", width=6, command=lambda: registrar_clique(nome, "zerar")).grid(row=linha, column=coluna_base + 2, padx=4)
        tk.Button(frame_botoes, text="-1", width=4, command=lambda: registrar_clique(nome, "subtrair")).grid(row=linha, column=coluna_base + 3, padx=4)
    else:
        tk.Label(frame_botoes, width=6, bg="black").grid(row=linha, column=coluna_base + 2)
        tk.Label(frame_botoes, width=6, bg="black").grid(row=linha, column=coluna_base + 3)

for idx, nome in enumerate(COLUNAS):
    criar_linha_botao(nome, idx)

frame_grafico = tk.Frame(main_frame, bg="black")
frame_grafico.pack(side=tk.LEFT, padx=20, anchor="n")

fig, ax = plt.subplots(figsize=(4.5, 4.5), facecolor='black')
canvas = FigureCanvasTkAgg(fig, master=frame_grafico)
canvas_widget = canvas.get_tk_widget()
canvas_widget.pack()

def atualizar_grafico():
    cliques = {k: v for k, v in dados_cache.items() if k != "Respondidos" and v > 0}
    ax.clear()
    if cliques:
        nomes = list(cliques.keys())
        valores = list(cliques.values())
        cores = [
            "#FF6347", "#FFD700", "#40E0D0", "#9370DB", "#3CB371",
            "#FFA07A", "#20B2AA", "#FF69B4", "#BA55D3", "#7B68EE",
            "#FF8C00", "#1E90FF", "#32CD32", "#DC143C", "#00CED1", "#FF1493"
        ]
        wedges, texts, autotexts = ax.pie(
            valores,
            labels=nomes,
            colors=cores[:len(nomes)],
            autopct='%1.1f%%',
            startangle=140,
            textprops={'color': 'white', 'fontsize': 9}
        )
        plt.setp(autotexts, size=9, weight="bold", color="white")
        ax.set_title("Distribuição de Cliques de Hoje", color="white", fontsize=12)
    canvas.draw()

dados_cache = obter_cliques_do_dia()
atualizar_labels()
after_id = root.after(15000, salvar_na_planilha)  # 15 segundos intervalo
agendar_zerar_dados_diariamente()
root.mainloop()
