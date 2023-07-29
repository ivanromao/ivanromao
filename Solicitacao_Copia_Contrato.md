# PARTE 1: SOLICITAÇÃO DE CÓPIA DE CONTRATO

import tkinter as tk
from tkinter import messagebox
import os
import openpyxl
from openpyxl.utils import get_column_letter
from datetime import datetime

# Função para criar a planilha com os dados fornecidos
def criar_planilha(prop_list, opcao, nome_solicitante):
    hoje = datetime.now().strftime("%Y-%m-%d")
    pasta_destino = "C:/Users/Usuário/Documents/Robô"
    if not os.path.exists(pasta_destino):
        os.makedirs(pasta_destino)

    nome_arquivo = f"{nome_solicitante}_{hoje}_{len(prop_list)}"
    caminho_excel = os.path.join(pasta_destino, f"{nome_arquivo}.xlsx")

    workbook = openpyxl.Workbook()
    workbook.create_sheet("CP", 0)
    workbook.create_sheet("BIOMETRIA", 1)

    cp_sheet = workbook["CP"]
    bio_sheet = workbook["BIOMETRIA"]

    cp_sheet["A1"] = "PROPOSTA"
    bio_sheet["A1"] = "PROPOSTA"

    row_cp = 2
    row_bio = 2
    for proposta in prop_list:
        if opcao.get() == "Normal":
            cp_sheet.cell(row=row_cp, column=1, value=proposta)
            row_cp += 1
        elif opcao.get() == "Bio":
            bio_sheet.cell(row=row_bio, column=1, value=proposta)
            row_bio += 1

    workbook.save(caminho_excel)
    messagebox.showinfo("Sucesso", f"Planilha criada e salva com o nome {nome_arquivo}.xlsx")

# Função chamada quando o botão "Enviar" é clicado
def on_click_enviar():
    prop_list = entry_proposta.get().split(",")
    nome_solicitante = entry_nome.get()

    if not prop_list or not nome_solicitante:
        messagebox.showwarning("Aviso", "Preencha todos os campos!")
        return

    criar_planilha(prop_list, opcao, nome_solicitante)

# Criar a janela de aplicativo
app = tk.Tk()
app.title("Mesalizador 🚀")
app.geometry("400x250")

# Criar widgets
label_instrucoes = tk.Label(app, text="Digite os números de proposta (separados por vírgula):")
label_instrucoes.pack(pady=10)

entry_proposta = tk.Entry(app, width=40)
entry_proposta.pack(pady=5)

opcao = tk.StringVar()  # Variável para armazenar a opção selecionada
opcao.set("")  # Valor inicial vazio

radio_normal = tk.Radiobutton(app, text="Normal", variable=opcao, value="Normal")
radio_normal.pack()
radio_bio = tk.Radiobutton(app, text="Bio", variable=opcao, value="Bio")
radio_bio.pack()

label_nome = tk.Label(app, text="Digite o seu nome:")
label_nome.pack(pady=10)

entry_nome = tk.Entry(app, width=40)
entry_nome.pack()

botao_enviar = tk.Button(app, text="Enviar", command=on_click_enviar)
botao_enviar.pack(pady=15)

# Iniciar o loop da aplicação
app.mainloop()
