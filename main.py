import joblib  # Para salvar e carregar o modelo
from concurrent.futures import ThreadPoolExecutor
from unidecode import unidecode
from selenium.webdriver.support.ui import Select
import threading
import subprocess
import os
import time
from tqdm import tqdm
import shutil
import json
from tqdm import tqdm
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
from selenium.common.exceptions import *
import re
import sys
import numpy as np
import cv2
import requests
from datetime import datetime

all_dados = pd.DataFrame()

start_row = 20  
end_row = 37
num_rows = end_row - start_row

df = pd.read_excel("GESTÃO DE AÇÕES E-COMMERCE.xlsx", usecols='C:O', skiprows=start_row, nrows=num_rows, engine='openpyxl', sheet_name="POLÍTICA COMERCIAL Dez24")

df.columns = ['PRODUTO', 'inutil1', 'SITE', 'COLUNA3','inutil2', 'CLÁSSICO ML', 'COLUNA5','inutil3', 'PREMIUM ML', 'COLUNA7','inutil4', 'MARKETPLACES', 'COLUNA9']


data_atual = datetime.now()
data_formatada = data_atual.strftime("%d/%m/%Y")

for index, i in df.iterrows():
    if i['PRODUTO'] == "FONTE 40A":
        fonte40Marketplace = round(i['COLUNA3'] , 2) ;
        fonte40Classico = round(i['COLUNA5'] , 2) ;
        fonte40Premium = round(i['COLUNA7'] , 2) ;
        fonte40PremiumPrice = round(i['PREMIUM ML'] , 2) ;
        fonte40ClassicoPrice = round(i['CLÁSSICO ML'] , 2) ;
        fonte40Marketplaceprice = round(i['SITE'] , 2) ;
    elif i['PRODUTO'] == "FONTE 60A":
        fonte60Marketplace = round(i['COLUNA3'] , 2) ;
        fonte60Classico = round(i['COLUNA5'] , 2) ;
        fonte60Premium = round(i['COLUNA7'] , 2) ;
        fonte60PremiumPrice = round(i['PREMIUM ML'] , 2) ;
        fonte60ClassicoPrice = round(i['CLÁSSICO ML'] , 2) ;
        fonte60Marketplaceprice = round(i['SITE'] , 2) ;
    elif i['PRODUTO'] == "FONTE 60A LITE":
        fonte60liteMarketplace = round(i['COLUNA3'] , 2) ;
        fonte60liteClassico = round(i['COLUNA5'] , 2) ;
        fonte60litePremium = round(i['COLUNA7'] , 2) ;
        fonte60litePremiumPrice = round(i['PREMIUM ML'] , 2) ;
        fonte60liteClassicoPrice = round(i['CLÁSSICO ML'] , 2) ;
        fonte60liteMarketplaceprice = round(i['SITE'] , 2) ;
    elif i['PRODUTO'] == "FONTE 70A":
        fonte70Marketplace = round(i['COLUNA3'] , 2) ;
        fonte70Classico = round(i['COLUNA5'] , 2) ;
        fonte70Premium = round(i['COLUNA7'] , 2) ;
        fonte70PremiumPrice = round(i['PREMIUM ML'] , 2) ;
        fonte70ClassicoPrice = round(i['CLÁSSICO ML'] , 2) ;
        fonte70Marketplaceprice = round(i['SITE'] , 2) ;
    elif i['PRODUTO'] == "FONTE 70A LITE":
        fonte70liteMarketplace = round(i['COLUNA3'] , 2) ;
        fonte70liteClassico = round(i['COLUNA5'] , 2) ;
        fonte70litePremium = round(i['COLUNA7'] , 2) ;
        fonte70litePremiumPrice = round(i['PREMIUM ML'] , 2) ;
        fonte70liteClassicoPrice = round(i['CLÁSSICO ML'] , 2) ;
        fonte70liteMarketplaceprice = round(i['SITE'] , 2) ;
    elif i['PRODUTO'] == "FONTE 90 BOB":
        fonte90bobMarketplace = round(i['COLUNA3'] , 2) ;
        fonte90bobClassico = round(i['COLUNA5'] , 2) ;
        fonte90bobPremium = round(i['COLUNA7'] , 2) ;
        fonte90bobPremiumPrice = round(i['PREMIUM ML'] , 2) ;
        fonte90bobClassicoPrice = round(i['CLÁSSICO ML'] , 2) ;
        fonte90bobMarketplaceprice = round(i['SITE'] , 2) ;
    elif i['PRODUTO'] == "FONTE 120 BOB":
        fonte120bobMarketplace = round(i['COLUNA3'] , 2) ;
        fonte120bobClassico = round(i['COLUNA5'] , 2) ;
        fonte120bobPremium = round(i['COLUNA7'] , 2) ;
        fonte120bobPremiumPrice = round(i['PREMIUM ML'] , 2) ;
        fonte120bobClassicoPrice = round(i['CLÁSSICO ML'] , 2) ;
        fonte120bobMarketplaceprice = round(i['SITE'] , 2) ;
    elif i['PRODUTO'] == "FONTE 120A LITE":
        fonte120liteMarketplace = round(i['COLUNA3'] , 2) ;
        fonte120liteClassico = round(i['COLUNA5'] , 2) ;
        fonte120litePremium = round(i['COLUNA7'] , 2) ;
        fonte120litePremiumPrice = round(i['PREMIUM ML'] , 2) ;
        fonte120liteClassicoPrice = round(i['CLÁSSICO ML'] , 2) ;
        fonte120liteMarketplaceprice = round(i['SITE'] , 2) ;
    elif i['PRODUTO'] == "FONTE 120A":
        fonte120Marketplace = round(i['COLUNA3'] , 2) ;
        fonte120Classico = round(i['COLUNA5'] , 2) ;
        fonte120Premium = round(i['COLUNA7'] , 2) ;
        fonte120PremiumPrice = round(i['PREMIUM ML'] , 2) ;
        fonte120ClassicoPrice = round(i['CLÁSSICO ML'] , 2) ;
        fonte120Marketplaceprice = round(i['SITE'] , 2) ;
    elif i['PRODUTO'] == "FONTE 200 BOB":
        fonte200bobMarketplace = round(i['COLUNA3'] , 2) ;
        fonte200bobClassico = round(i['COLUNA5'] , 2) ;
        fonte200bobPremium = round(i['COLUNA7'] , 2) ;
        fonte200bobPremiumPrice = round(i['PREMIUM ML'] , 2) ;
        fonte200bobClassicoPrice = round(i['CLÁSSICO ML'] , 2) ;
        fonte200bobMarketplaceprice = round(i['SITE'] , 2) ;
    elif i['PRODUTO'] == "FONTE 200A LITE":
        fonte200liteMarketplace = round(i['COLUNA3'] , 2) ;
        fonte200liteClassico = round(i['COLUNA5'] , 2) ;
        fonte200litePremium = round(i['COLUNA7'] , 2) ;
        fonte200litePremiumPrice = round(i['PREMIUM ML'] , 2) ;
        fonte200liteClassicoPrice = round(i['CLÁSSICO ML'] , 2) ;
        fonte200liteMarketplaceprice = round(i['SITE'] , 2) ;
    elif i['PRODUTO'] == "FONTE 200 MONO":
        fonte200monoMarketplace = round(i['COLUNA3'] , 2) ;
        fonte200monoClassico = round(i['COLUNA5'] , 2) ;
        fonte200monoPremium = round(i['COLUNA7'] , 2) ;
        fonte200monoPremiumPrice = round(i['PREMIUM ML'] , 2) ;
        fonte200monoClassicoPrice = round(i['CLÁSSICO ML'] , 2) ;
        fonte200monoMarketplaceprice = round(i['SITE'] , 2) ;
    elif i['PRODUTO'] == "FONTE 200A":
        fonte200Marketplace = round(i['COLUNA3'] , 2) ;
        fonte200Classico = round(i['COLUNA5'] , 2) ;
        fonte200Premium = round(i['COLUNA7'] , 2) ;
        fonte200PremiumPrice = round(i['PREMIUM ML'] , 2) ;
        fonte200ClassicoPrice = round(i['CLÁSSICO ML'] , 2) ;
        fonte200Marketplaceprice = round(i['SITE'] , 2) ;
    elif i['PRODUTO'] == "K1200":
        controleK1200Marketplace = round(i['COLUNA3'] , 2) ;
        controleK1200Classico = round(i['COLUNA5'] , 2) ;
        controleK1200Premium = round(i['COLUNA7'] , 2) ;
        controleK1200PremiumPrice = round(i['PREMIUM ML'] , 2) ;
        controleK1200ClassicoPrice = round(i['CLÁSSICO ML'] , 2) ;
        controleK1200Marketplaceprice = round(i['SITE'] , 2) ;
    elif i['PRODUTO'] == "K600":
        controleK600Marketplace = round(i['COLUNA3'] , 2) ;
        controleK600Classico = round(i['COLUNA5'] , 2) ;
        controleK600Premium = round(i['COLUNA7'] , 2) ;
        controleK600PremiumPrice = round(i['PREMIUM ML'] , 2) ;
        controleK600ClassicoPrice = round(i['CLÁSSICO ML'] , 2) ;
        controleK600Marketplaceprice = round(i['SITE'] , 2) ;
    elif i['PRODUTO'] == "CONTROLE WR":
        controleRedlineMarketplace = round(i['COLUNA3'] , 2) ;
        controleRedlineClassico = round(i['COLUNA5'] , 2) ;
        controleRedlinePremium = round(i['COLUNA7'] , 2) ;
        controleRedlinePremiumPrice = round(i['PREMIUM ML'] , 2) ;
        controleRedlineClassicoPrice = round(i['CLÁSSICO ML'] , 2) ;
        controleRedlineMarketplaceprice = round(i['SITE'] , 2) ;
    elif i['PRODUTO'] == "ACQUA":
        controleAcquaMarketplace = round(i['COLUNA3'] , 2) ;
        controleAcquaClassico = round(i['COLUNA5'] , 2) ;
        controleAcquaPremium = round(i['COLUNA7'] , 2) ;
        controleAcquaPremiumPrice = round(i['PREMIUM ML'] , 2) ;
        controleAcquaClassicoPrice = round(i['CLÁSSICO ML'] , 2) ;
        controleAcquaMarketplaceprice = round(i['SITE'] , 2) ;
        
        
def SelecionarFonte(item):
    price = item["Preço Unitário"]
    tipo = unidecode(item["tipo"].strip().lower())
    if item['modelo'] == "FONTE 40A":
        if tipo == "classico" and price > fonte40Classico:
            return f"FORA,{fonte40Classico}"
        elif tipo == "premium" and price > fonte40Premium:
            return f"FORA,{fonte40Premium}"
        if tipo == "classico" and price <fonte40Classico:
            return f"DENTRO,{fonte40Classico}"
        elif tipo == "premium" and price < fonte40Premium:
            return f"DENTRO,{fonte40Premium}"

    if item['modelo'] == "FONTE 60A":
        if tipo == "classico" and price > fonte60Classico:
            return f"FORA,{fonte60Classico}"
        elif tipo == "premium" and price > fonte60Premium:
            return f"FORA,{fonte60Premium}"
        if tipo == "classico" and price <fonte60Classico:
            return f"DENTRO,{fonte60Classico}"
        elif tipo == "premium" and price < fonte60Premium:
            return f"DENTRO,{fonte60Premium}"

    if item['modelo'] == "FONTE LITE 60A":
        if tipo == "classico" and price > fonte60liteClassico:
            return f"FORA,{fonte60liteClassico}"
        elif tipo == "premium" and price > fonte60litePremium:
            return f"FORA,{fonte60litePremium}"
        if tipo == "classico" and price <fonte60liteClassico:
            return f"DENTRO,{fonte60liteClassico}"
        elif tipo == "premium" and price < fonte60litePremium:
            return f"DENTRO,{fonte60litePremium}"

    if item['modelo'] == "FONTE 70A":
        if tipo == "classico" and price > fonte70Classico:
            return f"FORA,{fonte70Classico}"
        elif tipo == "premium" and price > fonte70Premium:
            return f"FORA,{fonte70Premium}"
        if tipo == "classico" and price <fonte70Classico:
            return f"DENTRO,{fonte70Classico}"
        elif tipo == "premium" and price < fonte70Premium:
            return f"DENTRO,{fonte70Premium}"

    if item['modelo'] == "FONTE LITE 70A":
        if tipo == "classico" and price > fonte70liteClassico:
            return f"FORA,{fonte70liteClassico}"
        elif tipo == "premium" and price > fonte70litePremium:
            return f"FORA,{fonte70litePremium}"
        if tipo == "classico" and price <fonte70liteClassico:
            return f"DENTRO,{fonte70liteClassico}"
        elif tipo == "premium" and price < fonte70litePremium:
            return f"DENTRO,{fonte70litePremium}"

    if item['modelo'] == "FONTE BOB 90A":
        if tipo == "classico" and price > fonte90bobClassico:
            return f"FORA,{fonte90bobClassico}"
        elif tipo == "premium" and price > fonte90bobPremium:
            return f"FORA,{fonte90bobPremium}"
        if tipo == "classico" and price <fonte90bobClassico:
            return f"DENTRO,{fonte90bobClassico}"
        elif tipo == "premium" and price < fonte90bobPremium:
            return f"DENTRO,{fonte90bobPremium}"

    if item['modelo'] == "FONTE 120A":
        if tipo == "classico" and price > fonte120Classico:
            return f"FORA,{fonte120Classico}"
        elif tipo == "premium" and price > fonte120Premium:
            return f"FORA,{fonte120Premium}"
        if tipo == "classico" and price <fonte120Classico:
            return f"DENTRO,{fonte120Classico}"
        elif tipo == "premium" and price < fonte120Premium:
            return f"DENTRO,{fonte120Premium}"

    if item['modelo'] == "FONTE LITE 120A":
        if tipo == "classico" and price > fonte120liteClassico:
            return f"FORA,{fonte120liteClassico}"
        elif tipo == "premium" and price > fonte120litePremium:
            return f"FORA,{fonte120litePremium}"
        if tipo == "classico" and price <fonte120liteClassico:
            return f"DENTRO,{fonte120liteClassico}"
        elif tipo == "premium" and price < fonte120litePremium:
            return f"DENTRO,{fonte120litePremium}"

    if item['modelo'] == "FONTE BOB 120A":
        if tipo == "classico" and price > fonte120bobClassico:
            return f"FORA,{fonte120bobClassico}"
        elif tipo == "premium" and price > fonte120bobPremium:
            return f"FORA,{fonte120bobPremium}"
        if tipo == "classico" and price <fonte120bobClassico:
            return f"DENTRO,{fonte120bobClassico}"
        elif tipo == "premium" and price < fonte120bobPremium:
            return f"DENTRO,{fonte120bobPremium}"

    if item['modelo'] == "FONTE 200A":
        if tipo == "classico" and price > fonte200Classico:
            return f"FORA,{fonte200Classico}"
        elif tipo == "premium" and price > fonte200Premium:
            return f"FORA,{fonte200Premium}"
        if tipo == "classico" and price <fonte200Classico:
            return f"DENTRO,{fonte200Classico}"
        elif tipo == "premium" and price < fonte200Premium:
            return f"DENTRO,{fonte200Premium}"

    if item['modelo'] == "FONTE MONO 200A":
        if tipo == "classico" and price > fonte200monoClassico:
            return f"FORA,{fonte200monoClassico}"
        elif tipo == "premium" and price > fonte200monoPremium:
            return f"FORA,{fonte200monoPremium}"
        if tipo == "classico" and price <fonte200monoClassico:
            return f"DENTRO,{fonte200monoClassico}"
        elif tipo == "premium" and price < fonte200monoPremium:
            return f"DENTRO,{fonte200monoPremium}"

    if item['modelo'] == "FONTE LITE 200A":
        if tipo == "classico" and price > fonte200liteClassico:
            return f"FORA,{fonte200liteClassico}"
        elif tipo == "premium" and price > fonte200litePremium:
            return f"FORA,{fonte200litePremium}"
        if tipo == "classico" and price <fonte200liteClassico:
            return f"DENTRO,{fonte200liteClassico}"
        elif tipo == "premium" and price < fonte200litePremium:
            return f"DENTRO,{fonte200litePremium}"

    if item['modelo'] == "FONTE BOB 200A":
        if tipo == "classico" and price > fonte200bobClassico:
            return f"FORA,{fonte200bobClassico}"
        elif tipo == "premium" and price > fonte200bobPremium:
            return f"FORA,{fonte200bobPremium}"
        if tipo == "classico" and price <fonte200bobClassico:
            return f"DENTRO,{fonte200bobClassico}"
        elif tipo == "premium" and price < fonte200bobPremium:
            return f"DENTRO,{fonte200bobPremium}"
        
    return "DENTRO,0"



service = Service()
options = webdriver.ChromeOptions()
titulo_arquivo = ""
# options.add_argument("--headless=new")

options.add_argument("--disable-gpu")
options.add_argument("--disable-extensions")
prefs = {"profile.managed_default_content_settings.images": 2}
options.add_experimental_option("prefs", prefs)


def limpar_pasta(caminho_pasta):
    # Verifica se o caminho é um diretório
    if not os.path.isdir(caminho_pasta):
        print(f'O caminho "{caminho_pasta}" não é um diretório válido.')
        return
    
    try:
        # Percorre todos os arquivos na pasta
        for nome_arquivo in os.listdir(caminho_pasta):
            caminho_completo = os.path.join(caminho_pasta, nome_arquivo)
            # Verifica se é um arquivo (não um diretório)
            if os.path.isfile(caminho_completo):
                # Remove o arquivo
                os.remove(caminho_completo)
                print(f'Arquivo "{nome_arquivo}" removido com sucesso.')
            else:
                print(f'O item "{nome_arquivo}" não é um arquivo.')

        print(f'Todos os arquivos em "{caminho_pasta}" FORAm removidos.')
    except Exception as e:
        print(f'Ocorreu um erro ao tentar limpar a pasta: {e}')

def excluir_arquivo(caminho_arquivo):
    # Verifica se o arquivo existe
    if os.path.exists(caminho_arquivo):
        try:
            # Remove o arquivo
            os.remove(caminho_arquivo)
            print(f'Arquivo "{caminho_arquivo}" removido com sucesso.')
        except Exception as e:
            print(f'Ocorreu um erro ao tentar excluir o arquivo: {e}')
    else:
        print(f'O arquivo "{caminho_arquivo}" não existe.')
        
excluir_arquivo("planilha_analise_anuncios.xlsx")        
limpar_pasta("dados")

def download_image_from_url(url, path):
    response = requests.get(url)
    img = cv2.imdecode(np.frombuffer(response.content, np.uint8), cv2.IMREAD_COLOR)
    if img is not None:
        cv2.imwrite(path, img)
        return path
    else:
        return None



driver = webdriver.Chrome(service=service, options=options)
try:
    driver.get("https://corp.shoppingdeprecos.com.br/login")
    counter = 0
    while True:
        test = driver.find_elements(By.XPATH, '//*[@id="email"]')
        if test:
            break
        else:
            counter += 1
            if counter > 20:
                break;
            time.sleep(0.5)
    driver.find_element(By.XPATH, '//*[@id="email"]').send_keys("loja@jfaeletronicos.com")
    driver.find_element(By.XPATH, '//*[@id="password"]').send_keys("922982PC")
    driver.find_element(By.XPATH, '//*[@id="btnLogin"]').click()
    print("Fez login")
except TimeoutException as e:
    print(f"Timeout ao tentar carregar a página ou encontrar um elemento: {e}")
except NoSuchElementException as e:
    print(f"Elemento não encontrado na página: {e}")
except WebDriverException as e:
    print(f"Erro no WebDriver: {e}")

time.sleep(3)
driver.get("https://corp.shoppingdeprecos.com.br/vendedores/busca")


time.sleep(3)
driver.find_element(By.XPATH, '//*[@id="txtTermo"]').send_keys("jfa")
time.sleep(2)
driver.find_element(By.XPATH, '//*[@id="cmbQtde"]').click()
time.sleep(1)
driver.find_element(By.XPATH, '//*[@id="cmbQtde"]/option[10]').click()
time.sleep(1)

time.sleep(1)
driver.execute_script("tabela(0);")
time.sleep(8)

commands = []
urls = []

for i in driver.find_elements(By.XPATH ,'//*[@id="table_result"]/tbody/tr'):
    url = i.find_element(By.XPATH, ".//td[3]/a")
    url = url.get_attribute("href")
    loja = i.find_element(By.XPATH, './/td[2]/a')
    commands.append(url)
        



driver.quit()

for i in commands:
    offset = 0
    limit = 50 
    total_results = float('inf')
    products = [] 
    while offset < total_results:
        # Adiciona offset e limit à URL da API
        response = requests.get(f"https://api.mercadolibre.com/sites/MLB/search?seller_id={i.split('_seller*id_')[1]}&offset={offset}&limit={limit}&q=jfa")
        response = response.json()
        # Obtém o número total de resultados na primeira iteração
        if offset == 0:
            total_results = response.get("paging", {}).get("total", 0)
            loja = response.get("seller", "").get("nickname", "")

        # Processa os itens da página atual
        for item in response.get("results", []):
            if "jfa" in item.get("title", "").lower():
                product = {
                    "data": data_formatada,
                    'loja': loja,
                    'Produto': item.get("title", ""),
                    "modelo": "",
                    'Preço Unitário': item.get("price", ""),
                    "politica": "",
                    'full': "FULL" if item.get("shipping", "").get("logistic_type", "") == "fulfillment" else "",
                    'tipo': "classico" if item.get("listing_type_id", "") == "gold_special" else "premium",
                    'link': item.get("permalink", ""),
                }
                products.append(product)

        # Incrementa o offset para a próxima página
        offset += limit

        # Sai do loop se não houver mais resultados
        if len(response.get("results", [])) < limit:
            break

    # Restante do processamento permanece o mesmo
    if len(products) != 0:
        # Cria DataFrame
        df_products = pd.DataFrame(products)

        # Carrega modelo e label encoder
        for index, item in df_products.iterrows():
            price = item['Preço Unitário']
            title = unidecode(item['Produto'].lower())
            if "controle" in title:
                item.loc['modelo'] = "OUTROS"
                df_products.loc[index, 'modelo'] = item['modelo']
                continue
            
            item.loc['modelo'] = "OUTROS"
            if "lite" not in title and "light" not in title  and "controle" not in title and "usina" not in title and ("jfa" in title or "fonte carregador" in title or "fonte automotiva" in title or "fonte e carregador" in title or "carregador de baterias" in title):
                if "40a" in title or "40" in title or "40 amperes" in title or "40amperes" in title or "36a" in title or "36" in title or "36 amperes" in title or "36amperes" in title:
                    item['modelo'] = "FONTE 40A"
                    df_products.loc[index, 'modelo'] = item['modelo']
                    continue
                
            if ("lite" in title or "light" in title) and "controle" not in title and "usina" not in title and ("jfa" in title or "fonte carregador" in title or "fonte automotiva" in title or "fonte e carregador" in title or "carregador de baterias" in title):
                if "50a" in title or "50" in title or "50 amperes" in title or "50amperes" in title:
                    item['modelo'] = "FONTE LITE 50A"
                    df_products.loc[index, 'modelo'] = item['modelo']
                    continue
                
            if ("lite" in title or "light" in title) and "controle" not in title and "usina" not in title and ("jfa" in title or "fonte carregador" in title or "fonte automotiva" in title or "fonte e carregador" in title or "carregador de baterias" in title):
                if "40a" in title or "40" in title or "40 amperes" in title or "40amperes" in title or "36a" in title or "36" in title or "36 amperes" in title or "36amperes" in title:
                    item['modelo'] = "FONTE LITE 40A"
                    df_products.loc[index, 'modelo'] = item['modelo']
                    continue

            if "lite" not in title and "light" not in title  and "controle" not in title and "usina" not in title and ("jfa" in title or "fonte carregador" in title or "fonte automotiva" in title or "fonte e carregador" in title or "carregador de baterias" in title):
                if "60a" in title or "60" in title or "60 amperes" in title or "60amperes" in title or "60 a" in title:
                    item['modelo'] = "FONTE 60A"
                    df_products.loc[index, 'modelo'] = item['modelo']
                    continue

            if ("lite" in title or "light" in title) and "controle" not in title and "usina" not in title and ("jfa" in title or "fonte carregador" in title or "fonte automotiva" in title or "fonte e carregador" in title or "carregador de baterias" in title):
                if "60a" in title or "60" in title or "60 amperes" in title or "60amperes" in title or "60 a" in title: 
                    item['modelo'] = "FONTE LITE 60A"
                    df_products.loc[index, 'modelo'] = item['modelo']
                    continue

            
            if  "lite" not in title and "light" not in title  and "controle" not in title and "usina" not in title and ("jfa" in title or "fonte carregador" in title or "fonte automotiva" in title or "fonte e carregador" in title or "carregador de baterias" in title):
                if "70a" in title or "70" in title or "70 amperes" in title or "70amperes" in title or "70 a" in title:
                    item['modelo'] = "FONTE 70A"
                    df_products.loc[index, 'modelo'] = item['modelo']
                    continue
                    
            if "bob" in title and "lite" not in title and "light" not in title  and "controle" not in title and "usina" not in title and ("jfa" in title or "fonte carregador" in title or "fonte automotiva" in title or "fonte e carregador" in title or "carregador de baterias" in title):
                if "90a" in title or "90" in title or "90 amperes" in title or "90amperes" in title or "90 a" in title:    
                    item['modelo'] = "FONTE BOB 90A"
                    df_products.loc[index, 'modelo'] = item['modelo']
                    continue


            if ("lite" in title or "light" in title) and "controle" not in title and "usina" not in title and ("jfa" in title or "fonte carregador" in title or "fonte automotiva" in title or "fonte e carregador" in title or "carregador de baterias" in title):
                if "70a" in title or "70" in title or "70 amperes" in title or "70amperes" in title or "70 a" in title:
                    item['modelo'] = "FONTE LITE 70A"
                    df_products.loc[index, 'modelo'] = item['modelo']
                    continue
                    
            if "bob" not in title and "lite" not in title and "light" not in title  and "controle" not in title and "usina" not in title and ("jfa" in title or "fonte carregador" in title or "fonte automotiva" in title or "fonte e carregador" in title or "carregador de baterias" in title):
                if "120a" in title or "120" in title or "120 amperes" in title or "120amperes" in title or "120 a" in title: 
                    item['modelo'] = "FONTE 120A"
                    df_products.loc[index, 'modelo'] = item['modelo']
                    continue

            if "bob" not in title and  ("lite" in title or "light" in title) and "controle" not in title and "usina" not in title and ("jfa" in title or "fonte carregador" in title or "fonte automotiva" in title or "fonte e carregador" in title or "carregador de baterias" in title):
                if "120a" in title or "120" in title or "120 amperes" in title or "120amperes" in title or "120 a" in title:
                    item['modelo'] = "FONTE LITE 120A"
                    df_products.loc[index, 'modelo'] = item['modelo']
                    continue

            if "bob" not in title and "lite" not in title and "light" not in title and "controle" not in title and 'mono' not in title and 'monovolt' not in title and "220v" not in title:
                if "200a" in title or "200 " in title or "200 amperes" in title or "200amperes" in title or "200 a" in title:
                    item['modelo'] = "FONTE 200A"
                    df_products.loc[index, 'modelo'] = item['modelo']
                    continue

            if "bob" not in title and  ("lite" in title or "light" in title) and "controle" not in title and 'mono' not in title and 'monovolt' not in title:
                if "200a" in title or "200" in title or "200 amperes" in title or "200amperes" in title or "200 a" in title:
                    item['modelo'] = "FONTE LITE 200A"
                    df_products.loc[index, 'modelo'] = item['modelo']
                    continue

            if "bob" not in title and  ("lite" in title or "light" in title) and "controle" not in title and 'mono' in title and 'monovolt' in title:
                if "200a" in title or "200" in title or "200 amperes" in title or "200amperes" in title or "200 a" in title:
                    item['modelo'] = "FONTE LITE 200A"
                    df_products.loc[index, 'modelo'] = item['modelo']
                    continue


            if "bob" in title and "lite" not in title and "light" not in title  and "controle" not in title and "usina" not in title and ("jfa" in title or "fonte carregador" in title or "fonte automotiva" in title or "fonte e carregador" in title or "carregador de baterias" in title):
                if "120a" in title or "120" in title or "120 amperes" in title or "120amperes" in title or "120 a" in title:
                    item['modelo'] = "FONTE BOB 120A"
                    df_products.loc[index, 'modelo'] = item['modelo']
                    continue
                    
            if "bob" in title and "lite" not in title and "light" not in title  and "controle" not in title and 'mono' not in title and 'mono' not in title and 'monovolt' not in title and "usina" not in title and ("jfa" in title or "fonte carregador" in title or "fonte automotiva" in title or "fonte e carregador" in title or "carregador de baterias" in title):
                if "200a" in title or "200" in title or "200 amperes" in title or "200amperes" in title or "200 a" in title:
                    item['modelo'] = "FONTE BOB 200A"
                    df_products.loc[index, 'modelo'] = item['modelo']
                    continue


            if "bob" not in title and "lite" not in title and "light" not in title  and "controle" not in title and ("mono" in title or "220v" in title or "monovolt" in title):
                if "200a" in title or "200" in title or "200 amperes" in title or "200amperes" in title or "200 a" in title:
                    item['modelo'] = "FONTE 200A MONO"
                    df_products.loc[index, 'modelo'] = item['modelo']
                    continue
                    
            df_products.loc[index, 'modelo'] = item['modelo']
        # Processa política
        for index, item in df_products.iterrows():
            fonte = SelecionarFonte(item).split(",")
            df_products.loc[index, 'politica'] = fonte[0]
        
        # Concatena com todos os dados
        all_dados = pd.concat([all_dados, df_products])

# Salva o arquivo final

all_dados.columns = ['data', 'loja', 'nome', 'modelo', 'preco', 'poltica', 'full', 'tipo', 'link']
all_dados.to_excel("planilha_analise_anuncios.xlsx", index=False)


# import pandas as pd
# import os.path

# from google.auth.transport.requests import Request
# from google.oauth2.credentials import Credentials
# from google_auth_oauthlib.flow import InstalledAppFlow
# from googleapiclient.discovery import build
# from googleapiclient.errors import HttpError

# SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]  # Read/write access
# SPREADSHEET_ID = "1vMwTxe1dkYZpp9ti1C8chyd5qY3bkff3jb9PqA2LfJg"
# DATA_RANGE = "DADOS!A1:I" 

# def get_sheets_credentials():
#     """Handles authentication with the Google Sheets API."""
#     creds = None
#     if os.path.exists("token.json"):
#         creds = Credentials.from_authorized_user_file("token.json", SCOPES)
#     if not creds or not creds.valid:
#         if creds and creds.expired and creds.refresh_token:
#             creds.refresh(Request())
#         else:
#             flow = InstalledAppFlow.from_client_secrets_file(
#                 "credentials.json", SCOPES
#             )
#             creds = flow.run_local_server(port=0)
#         with open("token.json", "w") as token:
#             token.write(creds.to_json())
#     return creds

# def get_next_empty_row(spreadsheet_id, range_name):
#     """Find the next empty row in the specified Google Sheet range."""
#     creds = get_sheets_credentials()
#     service = build("sheets", "v4", credentials=creds)
#     sheet = service.spreadsheets()

#     try:
#       result = sheet.values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
#       rows = result.get('values', [])
#       if not rows: # if no data
#           return 1 # start in the first row if there's no content
#       return len(rows) + 1
#     except HttpError as error:
#       print(f"An error occurred while getting data: {error}")
#       return 1 # start in the first row if there's an error


# def update_sheet_data(spreadsheet_id, range_name, values):
#     """Updates data in a specified range of a Google Sheet."""
#     creds = get_sheets_credentials()
#     service = build("sheets", "v4", credentials=creds)
#     sheet = service.spreadsheets()

#     try:
#         body = {"values": values}
#         result = (
#             sheet.values()
#             .update(
#                 spreadsheetId=spreadsheet_id,
#                 range=range_name,
#                 valueInputOption="USER_ENTERED",  # Important for direct data entry
#                 body=body,
#             )
#             .execute()
#         )
#         print(f"{result.get('updatedCells')} cells updated.")
#     except HttpError as error:
#         print(f"An error occurred while updating data: {error}")


# # Path to your Excel file
# file_path = os.path.join('', 'planilha_final.xlsx')  # Adjust path if necessary

# # --- Modified main execution block ---
# if __name__ == "__main__":

#     # 1. Read data from Excel file
#     try:
#         df = pd.read_excel(file_path)
#         # Clean your DataFrame
#         df = df.fillna('')  #Fill missing values with empty strings
#         for col in df.columns:
#             if df[col].dtype == 'object': # only clean string columns
#                 df[col] = df[col].astype(str).str.replace("'", "''", regex=False)
#     except FileNotFoundError:
#         print(f"Error: File not found at {file_path}")
#         exit()  # Exit if the file doesn't exist

#     # 2. Prepare data for Google Sheets update
#     data_to_update = []
#     header = df.columns.tolist()  # Include the header row
#     # data_to_update.append(header)  # Removed: header should be added only on the first time

#     for index, row in df.iterrows():
#         row_data = [row[col] for col in header]
#         data_to_update.append(row_data)  #Add each row from the dataframe


#     # 3. Determine next empty row and prepare the update range
#     next_row = get_next_empty_row(SPREADSHEET_ID, DATA_RANGE) # find the next empty row
#     update_range = f"DADOS!A{next_row}:I" # calculate the correct range to update based on the empty row

#     # 4. Update the Google Sheet
#     try:
#         # If it's the very first update of the sheet, add the header
#         if next_row == 1:
#             data_to_update.insert(0, header) # add header if it's the first row
#             update_sheet_data(SPREADSHEET_ID, update_range, data_to_update)
#         else:
#             update_sheet_data(SPREADSHEET_ID, update_range, data_to_update)
#         print("Data updated successfully.")

#     except Exception as e:
#         print(f"An error occurred during update: {e}")