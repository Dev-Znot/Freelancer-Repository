import xmltodict
import os
import pandas as pd

def pegar_info(nome_arquivo, valores):
    with open(f"nfs/{nome_arquivo}", "rb") as arquivo_xml:
        dic_arquivo = xmltodict.parse(arquivo_xml)
        if "nfeProc" in dic_arquivo:
            info_nf = dic_arquivo["nfeProc"]["NFe"]["infNFe"]
        else:
            info_nf = dic_arquivo["NFe"]["infNFe"]
        numero_nota = info_nf["@Id"]
        empresa_emissora = info_nf["emit"]["xNome"]
        nome_cliente = info_nf["dest"]["xNome"]
        endereço = info_nf["dest"]["enderDest"]
        peso = info_nf["transp"]["vol"]["pesoB"]
        valores.append([
            numero_nota,
            empresa_emissora,
            nome_cliente,
            endereço,
            peso
        ])
        


lista_arquivos = os.listdir('nfs')

colunas = ["numero_nota", "empresa_emissora", "nome_cliente", "endereço", "peso"]
valores = []

for arquivo in lista_arquivos:
    pegar_info(arquivo, valores)

tabela = pd.DataFrame(columns=colunas, data=valores)
tabela.to_excel("nfs.xlsx", index=False)