import tkinter as tk
import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta
from tkcalendar import DateEntry

janela = tk.Tk()
janela.title("FATURA DO CARTÃO")

janela.rowconfigure(0, weight=1)
janela.columnconfigure([0,1], weight=1)
mensagem1 = tk.Label(text="CONTROLE DE FATURA", fg="black", width=20, height=1, borderwidth=4, relief='solid')
mensagem1.grid(row=0,column=0, sticky="EW", columnspan=2, padx=5, pady=5)

mensagem2 = tk.Label(text="Digite o Valor:", fg="black", width=20, height=1) #bg="cor de fundo"
mensagem2.grid(row=1, column=0, padx=5, pady=5)

caixa_valor = tk.Entry(borderwidth=1, relief='solid')
caixa_valor.grid(row=1, column=1, padx=5, pady=5)

mensagem3 = tk.Label(text="Descrição:", fg="black", width=20, height=1)
mensagem3.grid(row=2, column=0, padx=5, pady=5)

descricao = tk.Entry(borderwidth=1, relief='solid')
descricao.grid(row=2, column=1, padx=5, pady=5)

mensagem_data = tk.Label(text="Data da compra:", fg="black", width=20, height=1)
mensagem_data.grid(row=3, column=0, padx=5, pady=5)

data = DateEntry(locale="pt_br")
data.grid(row=3, column=1, padx=5, pady=5)

mensagem_parcela = tk.Label(text="Qtd de parcelas:", fg="black", width=20, height=1)
mensagem_parcela.grid(row=4, column=0, padx=5, pady=5)

parcela = tk.Entry(borderwidth=1, relief='solid')
parcela.grid(row=4, column=1, padx=5, pady=5)
parcela.insert(0, "1")
tabela = pd.read_excel(r"C:\Users\BRUNO\PycharmProjects\sistemas\fatura.xlsx")

def leticia_():
    for j in range(int(parcela.get())):
        tabela = pd.read_excel(r"C:\Users\BRUNO\Desktop\FATURA\fatura.xlsx")
        i = len(tabela)
        tabela.loc[i, "DATA"] = datetime.strptime(data.get(), '%d/%m/%Y').date() + relativedelta(months=j)
        tabela.loc[i, "NOME"] = "Leticia"
        tabela.loc[i, "VALOR"] = float(caixa_valor.get().replace(",","."))/int(parcela.get())
        tabela.loc[i, "DESCRICAO"] = descricao.get()
        tabela.to_excel(r"C:\Users\BRUNO\Desktop\FATURA\fatura.xlsx", index=False)
        mensagem4 = tk.Label(text="Cadastrado com sucesso!", fg="black", width=20, height=1)
        mensagem4.grid(row=5, column=1)
        if j == int(parcela.get()) - 1:
            parcela.delete(0, 100)
            data.delete(0, 100)
            caixa_valor.delete(0, 100)
            descricao.delete(0, 100)
            parcela.insert(0, "1")

mensagem4 = tk.Label(text="", fg="black", width=20, height=1)
mensagem4.grid(row=5, column=1, padx=5, pady=5)

leticia = tk.Button(text="LETÍCIA", fg="white", bg="gray", width=20, height=1, command=leticia_)
leticia.grid(row=5, column=0, padx=5, pady=5)

def brenda_():
    for j in range(int(parcela.get())):
        tabela = pd.read_excel(r"C:\Users\BRUNO\Desktop\FATURA\fatura.xlsx")
        i = len(tabela)
        tabela.loc[i, "DATA"] = datetime.strptime(data.get(), '%d/%m/%Y').date() + relativedelta(months=j)
        tabela.loc[i, "NOME"] = "Brenda"
        tabela.loc[i, "VALOR"] = float(caixa_valor.get().replace(",", ".")) / int(parcela.get())
        tabela.loc[i, "DESCRICAO"] = descricao.get()
        tabela.to_excel(r"C:\Users\BRUNO\Desktop\FATURA\fatura.xlsx", index=False)
        mensagem5 = tk.Label(text="Cadastrado com sucesso!", fg="black", width=20, height=1)
        mensagem5.grid(row=6, column=1)
        if j == int(parcela.get()) - 1:
            parcela.delete(0, 100)
            data.delete(0, 100)
            caixa_valor.delete(0, 100)
            descricao.delete(0, 100)
            parcela.insert(0, "1")

mensagem5 = tk.Label(text="", fg="black", width=20, height=1)
mensagem5.grid(row=6, column=1, padx=5, pady=5)

brenda = tk.Button(text="BRENDA", fg="white", bg="gray", width=20, height=1, command=brenda_)
brenda.grid(row=6, column=0, padx=5, pady=5)

def camila_():
    for j in range(int(parcela.get())):
        tabela = pd.read_excel(r"C:\Users\BRUNO\Desktop\FATURA\fatura.xlsx")
        i = len(tabela)
        tabela.loc[i, "DATA"] = datetime.strptime(data.get(), '%d/%m/%Y').date() + relativedelta(months=j)
        tabela.loc[i, "NOME"] = "Camila"
        tabela.loc[i, "VALOR"] = float(caixa_valor.get().replace(",", ".")) / int(parcela.get())
        tabela.loc[i, "DESCRICAO"] = descricao.get()
        tabela.to_excel(r"C:\Users\BRUNO\Desktop\FATURA\fatura.xlsx", index=False)
        mensagem6 = tk.Label(text="Cadastrado com sucesso!", fg="black", width=20, height=1)
        mensagem6.grid(row=7, column=1)
        if j == int(parcela.get()) - 1:
            parcela.delete(0, 100)
            data.delete(0, 100)
            caixa_valor.delete(0, 100)
            descricao.delete(0, 100)
            parcela.insert(0, "1")

mensagem6 = tk.Label(text="", fg="black", width=20, height=1)
mensagem6.grid(row=7, column=1, padx=5, pady=5)

camila = tk.Button(text="CAMILA", fg="white", bg="gray", width=20, height=1, command=camila_)
camila.grid(row=7, column=0, padx=5, pady=5)

def walda_():
    for j in range(int(parcela.get())):
        tabela = pd.read_excel(r"C:\Users\BRUNO\Desktop\FATURA\fatura.xlsx")
        i = len(tabela)
        tabela.loc[i, "DATA"] = datetime.strptime(data.get(), '%d/%m/%Y').date() + relativedelta(months=j)
        tabela.loc[i, "NOME"] = "Walda"
        tabela.loc[i, "VALOR"] = float(caixa_valor.get().replace(",", ".")) / int(parcela.get())
        tabela.loc[i, "DESCRICAO"] = descricao.get()
        tabela.to_excel(r"C:\Users\BRUNO\Desktop\FATURA\fatura.xlsx", index=False)
        mensagem7 = tk.Label(text="Cadastrado com sucesso!", fg="black", width=20, height=1)
        mensagem7.grid(row=8, column=1)
        if j == int(parcela.get()) - 1:
            parcela.delete(0, 100)
            data.delete(0, 100)
            caixa_valor.delete(0, 100)
            descricao.delete(0, 100)
            parcela.insert(0, "1")

mensagem7 = tk.Label(text="", fg="black", width=20, height=1)
mensagem7.grid(row=8, column=1, padx=5, pady=5)

walda = tk.Button(text="WALDA", fg="white", bg="gray", width=20, height=1, command=walda_)
walda.grid(row=8, column=0, padx=5, pady=5)

def rozemira_():
    for j in range(int(parcela.get())):
        tabela = pd.read_excel(r"C:\Users\BRUNO\Desktop\FATURA\fatura.xlsx")
        i = len(tabela)
        tabela.loc[i, "DATA"] = datetime.strptime(data.get(), '%d/%m/%Y').date() + relativedelta(months=j)
        tabela.loc[i, "NOME"] = "Rozemira"
        tabela.loc[i, "VALOR"] = float(caixa_valor.get().replace(",", ".")) / int(parcela.get())
        tabela.loc[i, "DESCRICAO"] = descricao.get()
        tabela.to_excel(r"C:\Users\BRUNO\Desktop\FATURA\fatura.xlsx", index=False)
        mensagem8 = tk.Label(text="Cadastrado com sucesso!", fg="black", width=20, height=1)
        mensagem8.grid(row=9, column=1)
        if j == int(parcela.get()) - 1:
            parcela.delete(0, 100)
            data.delete(0, 100)
            caixa_valor.delete(0, 100)
            descricao.delete(0, 100)
            parcela.insert(0, "1")

mensagem8 = tk.Label(text="", fg="black", width=20, height=1)
mensagem8.grid(row=9, column=1, padx=5, pady=5)

rozemira = tk.Button(text="ROZEMIRA", fg="white", bg="gray", width=20, height=1, command=rozemira_)
rozemira.grid(row=9, column=0, padx=5, pady=5)

def outro_():
    for j in range(int(parcela.get())):
        tabela = pd.read_excel(r"C:\Users\BRUNO\Desktop\FATURA\fatura.xlsx")
        i = len(tabela)
        tabela.loc[i, "DATA"] = datetime.strptime(data.get(), '%d/%m/%Y').date() + relativedelta(months=j)
        tabela.loc[i, "NOME"] = "Outro"
        tabela.loc[i, "VALOR"] = float(caixa_valor.get().replace(",", ".")) / int(parcela.get())
        tabela.loc[i, "DESCRICAO"] = descricao.get()
        tabela.to_excel(r"C:\Users\BRUNO\Desktop\FATURA\fatura.xlsx", index=False)
        mensagem9 = tk.Label(text="Cadastrado com sucesso!", fg="black", width=20, height=1)
        mensagem9.grid(row=10, column=1)
        if j == int(parcela.get()) - 1:
            parcela.delete(0,100)
            data.delete(0, 100)
            caixa_valor.delete(0, 100)
            descricao.delete(0, 100)
            parcela.insert(0, "1")

mensagem9 = tk.Label(text="", fg="black", width=20, height=1)
mensagem9.grid(row=10, column=1, padx=5, pady=5)

outro = tk.Button(text="OUTRO", fg="white", bg="gray", width=20, height=1, command=outro_)
outro.grid(row=10, column=0, padx=5, pady=5)

janela.mainloop()