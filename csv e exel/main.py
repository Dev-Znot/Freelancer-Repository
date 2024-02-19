import os
from datetime import datetime
import pandas as pd
import win32com.client as win32

caminho = "base"
lista_arquivos = os.listdir(caminho)

tabela_consolidada = pd.DataFrame()

for nome_arquivo in lista_arquivos:
    tabela_vendas = pd.read_csv(os.path.join(caminho, nome_arquivo))
    tabela_vendas["Data de Venda"] = pd.to_datetime("01/01/1900") + pd.to_timedelta(tabela_vendas("Data de Venda"), unit="d")

    #Concatena
    tabela_consolidada = pd.concat([tabela_consolidada, tabela_vendas])

# Ordenando pela coluna "Data de Venda"
tabela_consolidada = tabela_consolidada.sort_values(by="Data de Venda")
# Resetando o indice da tabela
tabela_consolidada = tabela_consolidada.reset_index(drop=True)
tabela_consolidada.to_excel("Vendas.xlsx",index=False)


outlook = win32.Dispatch('outlook.application')
email = outlook.CreateItem(0)
email.To = 'username@outlook.com'
data_hoje = datetime.today().strftime("%d-%m-%Y")
email.Subject = f"Relatorio de Vendas {data_hoje}"
email.Body = f"""
Prezado,

Segue em anexo o Relatorio de vendas de {data_hoje} atualizado.
abs,
Programador
"""

caminho = os.getcwd()
anexo = os.path.join(caminho, "Vendas.xlsx")
email.Attachments.Add(anexo)

email.Send()

