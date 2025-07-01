# estudos de automação python com excel

import requests
from bs4 import BeautifulSoup
import pandas as pd
import time

# Lista de empresas da bolsa de valores para gerar a planilha automaticamente 
empresas = [
    ("Petrobras", "PETR4"),
    ("Vale", "VALE3"),
    ("Itaú", "ITUB4"),
    ("Bradesco", "BBDC4"),
    ("Magazine Luiza", "MGLU3"),
    ("Ambev", "ABEV3"),
    ("Banco do Brasil", "BBAS3"),
    ("Weg", "WEG3"),
    ("Ibovespa", "IBOV"),

]

# Criar a planilha se ela não existir 
df_empresas = pd.DataFrame(empresas, columns=["Nome da Empresa", "Código da Empresa"]) # df vem de dataframe, estrutura utilizada pela biblioteca pandas
df_empresas.to_excel("empresas_info.xlsx", index=False) # Função que cria a planilha 

# Lê a planilha com as empresas 
def_empresas = pd.read_excel("empresas_info.xlsx") # Função para ler o excel 

# URL base do Google Finance 
url_base = 'https:www.google.com/finance/quote/' # Cuidado para não errar a URL

# Lista para armazenar cotações 
cotacoes = []

# Loop para obter as cotações
for _, row in df_empresas.iterrows(): # Para cada linha dentro do nosso excel teremos o odf_empresas
    codigo_da_empresa = row["Código da Empresa"] # Função responsavel por buscar o código da empresa e jogar no google para achar a cotação

    # Verifica se é o índice Ibovespa 
    if codigo_da_empresa == "IBOV" :
        url = url_base + "IBOV:INDEXBVMF"
    else:
        url = url_base + f"{codigo_da_empresa}:BVMF"

# Estrutura de tentativa e erro 
    try: # tenta mostrar isso daqui
        response = requests.get(url, timeout=10) # Esturtura que entra no site do google, requisição e resposta
        response.raise_for_status()
    except requests.RequestException as e: # se não der e der erro, mostra isso daqui
        print(f"Erro ao acessar {codigo_da_empresa}: {e}")
        cotacoes.append("Erro")
        continue

    # Estrutura que faz com que o python no código lá do html 
    soup = BeautifulSoup(response.text, "html.perser")

    try: 
        cotacao_str = soup.find("div", class_="YM1Kec fxKbKc").text.strip() # Vai buscar a div class e formatar esse texto
        cotacoes.append(cotacao_str) # Mantém em formato R$ 12,34 e joga na lista geral
    except AttributeError:
        print(f"Não foi possivel encontrar a cotação de {row["Nome da Empresa"]}")
        cotacoes.append("Não disponível")

    time.sleep(1) #Evita bloqueios no site por muitas requisições 

# Adiciona as cotações no DataFrame 
df_empresas["Cotações"] = cotacoes

# Salva o novo arquivo no Excel 
df_empresas.to_excel("cotacoes_atualizadas.xlsx", index=False)

print("Planilha 'cotacoes_atualizadas.xlsx' atualizada com sucesso!")

