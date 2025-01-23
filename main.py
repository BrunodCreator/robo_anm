from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from time import sleep
import os
from func.selecionar_combobox import selecionar_subs_agrupadora, selecionar_substancia, selecionar_regiao, selecionar_estado, selecionar_municipio
from func.selecionar_radio import selecionar_radio_empresa, selecionar_radio_operacao
from func.clicar_gera import clicar_gera

#configurando o webdriver

service = Service(ChromeDriverManager().install())
navegador = webdriver.Chrome(service=service)


url = 'https://sistemas.anm.gov.br/arrecadacao/extra/relatorios/cfem/maiores_arrecadadores.aspx'
navegador.get(url)

sleep(15)
def capturar_primeiras_seis_colunas():
    # Captura o elemento com o seletor CSS
    elemento = navegador.find_element(By.CSS_SELECTOR, '#ctl00_ContentPlaceHolder1_dvResultado > table.tabelaRelatorio')

    # Extrai o texto do elemento
    texto = elemento.text
    print(f"Dado capturado: {texto}")

    # Transforma o texto em um dicionário
    dados = {}
    for linha in texto.split("\n"):  # Divide o texto em linhas
        if " : " in linha:  # Verifica se a linha contém " : "
            chave, valor = linha.split(" : ", 1)  # Divide a linha em chave e valor
            dados[chave.strip()] = valor.strip()  # Adiciona ao dicionário removendo espaços extras

    print(dados)
    # ano = dados['Ano']
    # subs_agrupadora = dados['Substância Agrupadora']
    # substancia = dados['Substância']
    # regiao = dados['Regiao']
    # estado = dados['Estado']
    # municipio = dados['Municipio']
    # # Exibe o dicionário
    # return ano, subs_agrupadora, substancia, regiao, estado, municipio

# def capturar_coluna_inferior():
    
capturar_primeiras_seis_colunas()