import pandas as pd  
import numpy as np
import yfinance as yf
from datetime import datetime
import tkinter as tk
from tkinter import filedialog

# Variável global para armazenar o caminho do arquivo selecionado
caminho_arquivo_excel = ""

# Tkinter para gerar a interface gráfica e selecionar o arquivo excel
def selecionar_arquivo():
    global caminho_arquivo_excel
    arquivo = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx;*.xlsm")])
    if arquivo:
        print("Arquivo selecionado:", arquivo)
        caminho_arquivo_excel = arquivo
        # Fecha a janela após selecionar o arquivo
        janela.destroy()

janela = tk.Tk()
janela.title("Selecionar Arquivo Excel")
janela.geometry("300x100")

btn_selecionar = tk.Button(janela, text="Selecionar Arquivo Excel", command=selecionar_arquivo)
btn_selecionar.pack(pady=20)

janela.mainloop()


# Importar cotação LFTS11 do Yahoo Finance
ticker = "LFTS11.SA"  

dados = yf.download(ticker, start= datetime.today())

cotacao_final = dados['Close']

cotacao_final = float(cotacao_final)


 # Calculo para compra de cotas
df = pd.read_excel(caminho_arquivo_excel)

saldos_positivos = df[df.iloc[:, 1] > cotacao_final]

calculo_compra_cotas = (saldos_positivos.iloc[:, 1] / cotacao_final).astype(int)

#Calculo para venda de cotas

df = pd.read_excel(caminho_arquivo_excel)

saldos_negativos = df[df.iloc[:, 1] < 0]

calculo_venda_cotas = np.ceil(-saldos_negativos.iloc[:, 1] / cotacao_final).astype(int)



# Criar planilha TWAP e no modelo requisitado

cabecalho = ["ORDEM", "LADO", "ATIVO", "QUANTIDADE", "INICIO", "FIM", "PREÇO LIMITE"]
df_twap = pd.DataFrame(columns=cabecalho)

for index, row in saldos_positivos.iterrows():
    df_twap.loc[len(df_twap)] = ['TWAP', 'COMPRA', 'LFTS11', calculo_compra_cotas[index], '', '16:45', '']

for index, row in saldos_negativos.iterrows():
    df_twap.loc[len(df_twap)] = ['TWAP', 'VENDA', 'LFTS11', calculo_venda_cotas[index], '', '16:45', '']

nome_novo_arquivo = "modelo_TWAP.xlsx"
with pd.ExcelWriter(nome_novo_arquivo, engine='xlsxwriter') as writer:
    df_twap.to_excel(writer, index=False)

    workbook = writer.book
    worksheet = writer.sheets["Sheet1"] 

    formato_palavras_cabecalho = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#F2F2F2'})
    formato_palavras = workbook.add_format({'align': 'center'})

    for col_num, value in enumerate(cabecalho):
        worksheet.write(0, col_num, value, formato_palavras_cabecalho)

    for row_num in range(1, len(df_twap) + 1):
        for col_num in range(7): 
            if col_num != 3:  
                worksheet.write(row_num, col_num, df_twap.iloc[row_num - 1, col_num], formato_palavras)

    for i, width in enumerate(df_twap.drop(columns=["QUANTIDADE"]).map(lambda x: len(str(x))).max()):
        worksheet.set_column(i, i, max(width, len(cabecalho[i])) + 2)  




df_filtro = pd.read_excel(caminho_arquivo_excel)

df_filtro1 = df_filtro[df_filtro["Vl. Total"] >= cotacao_final]

df_filtro2 = df_filtro[df_filtro["Vl. Total"] < 0]

# Criar planilha Mesa no padrão requisitado

cabecalho = ["C\V", "Código", "Nome", "Qtd", "Qtd Aberta", "Preço", "Cliente", "Nome Cliente", "Agente de Custódia", "Conta de Custódia", "Código da Carteira", "N°"]
df_twap = pd.DataFrame(columns=cabecalho)

numeral_counter = 1

for index, row in saldos_positivos.iterrows():
    df_twap.loc[len(df_twap)] = ['C', 'LFTS11', '', calculo_compra_cotas[index], '', 'MERCADO', row["Cod. Conta Local"], "", '', '', '', numeral_counter]
    numeral_counter += 1

for index, row in saldos_negativos.iterrows():
    df_twap.loc[len(df_twap)] = ['V', 'LFTS11', '', calculo_venda_cotas[index], '', 'MERCADO', row["Cod. Conta Local"], "", '', '', '', numeral_counter]
    numeral_counter += 1

nome_novo_arquivo1 = "modelo_Mesa.xlsx"
with pd.ExcelWriter(nome_novo_arquivo1, engine='xlsxwriter') as writer:
    df_twap.to_excel(writer, index=False)

    workbook = writer.book
    worksheet = writer.sheets["Sheet1"]  

    
    worksheet.autofilter(0, 0, len(df_twap), len(cabecalho) - 1)

    formato_palavras_cabecalho2 = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#F2F2F2'})
    formato_palavras_cabecalho = workbook.add_format({'align': 'center'})

    for col_num, value in enumerate(cabecalho):
        worksheet.write(0, col_num, value, formato_palavras_cabecalho2)
