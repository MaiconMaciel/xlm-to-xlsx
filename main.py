import xmltodict
import os
import pandas as pd


def get_info(archive, valores):
    with open(f'nfs/{archive}', "rb") as xml_file:
        xml_dict = xmltodict.parse(xml_file)
        if "NFe" in xml_dict:
            info_nf = xml_dict["NFe"]["infNFe"]
            endereco = info_nf["entrega"]
        else:
            info_nf = xml_dict["nfeProc"]["NFe"]["infNFe"]
            endereco = info_nf["emit"]["enderEmit"]

        numero_nota = info_nf["@Id"]
        empresa_emissora = info_nf["emit"]["xNome"]
        nome_cliente = info_nf["dest"]["xNome"]
        if "vol" in info_nf["transp"]:
            peso = info_nf["transp"]["vol"]["pesoB"]
        else:
            peso = "Não Informado"

        valores.append([numero_nota, empresa_emissora, nome_cliente, endereco, peso])


arquivos = os.listdir("nfs")

colunas = ["numero_nota", "empresa_emissora", "nome_cliente", "endereco", "peso"]
valores = []

for arq in arquivos:
    get_info(arq, valores)

tabela = pd.DataFrame(columns=colunas, data=valores)
tabela.to_excel("NotasFiscais.xlsx", index=False)


