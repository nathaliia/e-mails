import win32com.client
import pandas as pd
import tkinter as tk
from tkinter import ttk
from collections import Counter
from datetime import datetime, timedelta
import os

# ========== 1. Regras de categorização ==========

regras = {
    'Financeiro': ['nfe_br@bionexo.com', 'tesouraria_br@bionexo.com', 'nmartins.dsrh@out.bionexo.com'],
}
assunto_regras = {
    'Folha Mensal': ['folha', 'mensal'],
    'Rescisão': ['rescisão', 'demissão'],
    'Férias': ['férias'],
}

# ========== 2. Conectar ao Outlook ==========

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
caixa_entrada = outlook.GetDefaultFolder(6)  # 6 = Caixa de Entrada
emails = caixa_entrada.Items
emails.Sort("[ReceivedTime]", True)  # ✅ Verificações iniciais

# Exibindo no terminal
print("📬 Total de e-mails na Caixa de Entrada:", len(emails))

# Exibindo os 10 primeiros e-mails
print("🔍 Exibindo os 10 primeiros e-mails:")
for email in list(emails)[:10]:
    if email.Class == 43:
        print("Data de Recebimento:", email.ReceivedTime.strftime('%d/%m/%Y %H:%M'))
        print("Assunto:", email.Subject)
        print("Remetente:", email.SenderEmailAddress)
        print("-" * 40)

# ========== 3. Processar os e-mails ==========

dados = []
referencias_emails = []

nao_lidos = 0
emails_ultimos_7_dias = 0
dados_por_categoria = {categoria: 0 for categoria in regras.keys()}

for email in list(emails)[:50]:
    if email.Class == 43:  # Apenas e-mails (MailItem)
        assunto = email.Subject or ""
        remetente = email.SenderEmailAddress or ""
        corpo = email.Body or ""
        lido = email.Unread

        # Contagem de e-mails não lidos
        if lido:
            nao_lidos += 1

        # Contagem de e-mails recebidos nos últimos 7 dias
        if email.ReceivedTime >= (datetime.now() - timedelta(days=7)):
            emails_ultimos_7_dias += 1

        # Verificação de status
        replies = email.ReplyRecipients
        foi_respondido = email.Sent or email.SenderEmailAddress in [r.Address for r in replies]

        if foi_respondido:
            status = "Retorno Enviado"
        elif "resposta" in corpo.lower() or "retorno" in corpo.lower():
            status = "Retorno"
        else:
            status = "Novo E-mail"

        # Categorização por remetente
        categoria = ""
        for nome_cat, lista_emails in regras.items():
            if remetente.lower() in lista_emails:
                categoria = nome_cat
                dados_por_categoria[nome_cat] += 1  # Incrementa a categoria

        # Categorização por palavras-chave no assunto
        for assunto_nome, palavras in assunto_regras.items():
            if any(p.lower() in assunto.lower() for p in palavras):
                categoria = assunto_nome

        dados.append([assunto, remetente, status, categoria])
        referencias_emails.append(email)

# ========== 4. Criar DataFrame ==========

df = pd.DataFrame(dados, columns=["Assunto", "Remetente", "Status", "Categoria"])

# Contagem de e-mails com o mesmo assunto
df['Contagem'] = df.groupby('Assunto')['Assunto'].transform('count')

# Remover duplicatas de e-mails (manter uma linha por remetente e assunto)
df = df.drop_duplicates(subset=['Assunto', 'Remetente'])

# ========== 5. Interface com Tkinter ==========

root = tk.Tk()
root.title("📬 Dashboard de E-mails - Outlook")
root.geometry("1000x600")

# Título
label_titulo = tk.Label(root, text="📬 Gerenciador de E-mails - Outlook", font=("Arial", 16, "bold"))
label_titulo.pack(pady=10)

# Tabela
tree = ttk.Treeview(root, columns=list(df.columns), show="headings", height=15)
for col in df.columns:
    tree.heading(col, text=col)
    tree.column(col, width=230)
for i, row in enumerate(df.values):
    tree.insert("", tk.END, values=list(row), tags=(f"row{i}",))
tree.pack(pady=10)

# Contadores principais
total = len(df)
respondidos = len(df[df["Status"] == "Retorno Enviado"])
pendentes = len(df[df["Status"] == "Novo E-mail"])
retornos = len(df[df["Status"] == "Retorno"])

label_info = tk.Label(
    root,
    text=f"📩 Recebidos: {total} | 📤 Respondidos: {respondidos} | ⏳ Pendentes: {pendentes} | 🔁 Retornos: {retornos}",
    font=("Arial", 12)
)
label_info.pack(pady=5)

# Exibir informações adicionais (não lidos, últimos 7 dias, e-mails por categoria)
label_info_adicionais = tk.Label(
    root,
    text=f"📑 Não Lidos: {nao_lidos} | 📅 E-mails nos últimos 7 dias: {emails_ultimos_7_dias}",
    font=("Arial", 12)
)
label_info_adicionais.pack(pady=5)

frame_categorias = tk.Frame(root)
frame_categorias.pack(pady=5)

contagem_categorias = Counter(df["Categoria"])
for categoria, qtd in contagem_categorias.items():
    if categoria:
        cor = "#A9DDF3"  # Azul claro
        tk.Label(
            frame_categorias,
            text=f"{qtd}\n{categoria}",
            bg=cor,
            font=("Arial", 10, "bold"),
            width=14,
            height=3,
            relief="raised"
        ).pack(side="left", padx=5)

# ========== 6. Evento de clique duplo para abrir o e-mail ==========

def abrir_email(index):
    try:
        referencias_emails[index].Display()
    except Exception as e:
        print(f"Erro ao abrir e-mail: {e}")

def on_double_click(event):
    selected_item = tree.selection()
    if selected_item:
        index = tree.index(selected_item[0])
        abrir_email(index)

tree.bind("<Double-1>", on_double_click)

# ========== 7. Atualizar ou Criar o arquivo Excel ==========

diretorio_atual = os.path.dirname(os.path.abspath(__file__))  # Diretório do script
nome_arquivo = os.path.join(diretorio_atual, "dashboard_emails.xlsx")

# Adicionar informações adicionais antes da lista de e-mails
info_adicionais = {
    'Não Lidos': nao_lidos,
    'E-mails Últimos 7 Dias': emails_ultimos_7_dias,
}

# Adicionar contagem de e-mails por categoria
for categoria, qtd in dados_por_categoria.items():
    info_adicionais[categoria] = qtd

try:
    # Se o arquivo já existir, carrega os dados existentes
    df_existente = pd.read_excel(nome_arquivo)

    # Concatenar os novos dados com os dados existentes
    df_final = pd.concat([df_existente, df]).drop_duplicates(subset=['Assunto', 'Remetente'])

    # Salvar o arquivo atualizado
    with pd.ExcelWriter(nome_arquivo, engine='openpyxl') as writer:
        # Criar uma nova planilha para informações adicionais
        pd.DataFrame([info_adicionais]).to_excel(writer, sheet_name='Info', index=False)
        # Criar ou atualizar a planilha com os dados dos e-mails
        df_final.to_excel(writer, sheet_name='E-mails', index=False)

    print(f"📁 Dados exportados automaticamente para: {nome_arquivo}")
except FileNotFoundError:
    # Se o arquivo não existir, cria um novo
    with pd.ExcelWriter(nome_arquivo, engine='openpyxl') as writer:
        # Criar uma nova planilha para informações adicionais
        pd.DataFrame([info_adicionais]).to_excel(writer, sheet_name='Info', index=False)
        # Criar a planilha com os dados dos e-mails
        df.to_excel(writer, sheet_name='E-mails', index=False)
    print(f"📁 Arquivo criado: {nome_arquivo}")

# Iniciar interface
root.mainloop()
