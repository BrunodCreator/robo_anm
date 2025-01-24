from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from openpyxl import workbook, load_workbook
import os
from datetime import datetime

# import preencher_formulario
from time import sleep

def capturar_dados_table():
    ano = datetime.now().year

    
    
caminho_planilha = os.path.join(os.path.expanduser("~"), "Desktop", "Arrecadadores CFEM.xlsx")

workbook = load_workbook(caminho_planilha)
sheet = workbook["Folha1"]


sheet['A1'] = 'Ano'
sheet['B1'] = 'Subs.Agrupadora'
sheet['C1'] = 'Substância'
sheet['D1'] = 'Região'
sheet['E1'] = 'Estado'
sheet['F1'] = 'Municipio'
sheet['G1'] = 'Arrecadador (Empresa)'
sheet['H1'] = 'Qtde Títulos'
sheet['K1'] = 'Operação'
sheet['I1'] = 'Recolhimento CFEM'
sheet['J1'] = '% Recolhimento CFEM'

linha_excel = 2

def capturar_primeiras_seis_colunas(func_navegador):
    # Captura o elemento com o seletor CSS
    elemento = func_navegador.find_element(By.CSS_SELECTOR, '#ctl00_ContentPlaceHolder1_dvResultado > table.tabelaRelatorio')

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



def capturar_dados_segunda_planilha(func_navegador):
    # Localize a tabela
    tabela = func_navegador.find_element(By.CSS_SELECTOR, '#ctl00_ContentPlaceHolder1_dvResultado > table.tabelaRelatorio')

    # Localize as linhas da tabela (exceto o cabeçalho)
    linhas = tabela.find_elements(By.CSS_SELECTOR, "tbody > tr")

    # Inicializa uma lista para armazenar os dados
    dados_tabela = []

    # Itera sobre as linhas e captura os dados das células
    for linha in linhas:
        colunas = linha.find_elements(By.TAG_NAME, "td")  # Captura todas as colunas da linha

        # Verifica se a linha tem pelo menos 6 colunas
        if len(colunas) >= 6:
            # Adiciona os dados da linha como um dicionário
            dados_tabela.append({
                "Arrecadador (Empresa)": colunas[1].text,  # Nome da empresa (coluna 2)
                "Qtde Títulos": colunas[2].text,          # Quantidade de títulos (coluna 3)
                "Operação": colunas[3].text,              # Valor de operação (coluna 4)
                "Recolhimento CFEM": colunas[4].text,     # Valor do recolhimento CFEM (coluna 5)
                "% Recolhimento CFEM": colunas[5].text    # Percentual de recolhimento CFEM (coluna 6)
            })
        else:
            # Caso a linha não tenha o número esperado de colunas, registre para depuração
            print(f"Linha ignorada por não ter colunas suficientes: {linha.text}")

    # Exibe os dados capturados
    for dado in dados_tabela:
        print(dado)
    return dados_tabela
        
def salvar_dados_planilha(dados_tabela, nome_arquivo="teste.xlsx"):
    # Caminho para salvar o arquivo na área de trabalho
    caminho_arquivo = os.path.join(os.path.expanduser("~"), "Desktop", 'teste.xlsx')

    # Verifica se a planilha já existe
    if os.path.exists(caminho_arquivo):
        # Abre a planilha existente
        workbook = load_workbook(caminho_arquivo)
    else:
        # Cria uma nova planilha
        workbook = Workbook()
    
    # Seleciona a primeira folha da planilha
    folha = workbook.active
    folha.title = "Arrecadadores"

    # Adiciona o cabeçalho (apenas se a planilha for nova)
    if folha.max_row == 1:
        folha.append(["Arrecadador (Empresa)", "Qtde Títulos", "Operação", "Recolhimento CFEM", "% Recolhimento CFEM"])

    # Adiciona os dados capturados
    for dado in dados_tabela:
        folha.append([
            dado["Arrecadador (Empresa)"],
            dado["Qtde Títulos"],
            dado["Operação"],
            dado["Recolhimento CFEM"],
            dado["% Recolhimento CFEM"]
        ])

    # Salva a planilha
    workbook.save(caminho_arquivo)
    print(f"Dados salvos na planilha '{caminho_arquivo}' com sucesso!")
    
    
    
    
def capturar_todos_os_dados(func_navegador):
    # Captura os dados da primeira tabela
    elemento = func_navegador.find_element(By.CSS_SELECTOR, '#ctl00_ContentPlaceHolder1_dvResultado > table.tabelaFormulario')
    texto = elemento.text

    # Transforma os dados da primeira tabela em um dicionário
    dados_gerais = {}
    for linha in texto.split("\n"):
        if " : " in linha:
            chave, valor = linha.split(" : ", 1)
            dados_gerais[chave.strip()] = valor.strip()

    print("Dados gerais capturados:", dados_gerais)

    # Captura os dados da segunda tabela
    tabela = func_navegador.find_element(By.CSS_SELECTOR, '#ctl00_ContentPlaceHolder1_dvResultado > table.tabelaRelatorio')
    linhas = tabela.find_elements(By.CSS_SELECTOR, "tbody > tr")

    # Combina os dados das duas tabelas em uma lista de dicionários
    dados_completos = []
    for linha in linhas:
        colunas = linha.find_elements(By.TAG_NAME, "td")
        if len(colunas) >= 6:
            # Cria um dicionário com os dados da linha
            linha_dados = {
                "Arrecadador (Empresa)": colunas[1].text,
                "Qtde Títulos": colunas[2].text,
                "Operação": colunas[3].text,
                "Recolhimento CFEM": colunas[4].text,
                "% Recolhimento CFEM": colunas[5].text
            }

            # Adiciona os dados gerais (primeira tabela) a cada linha
            linha_dados.update(dados_gerais)
            dados_completos.append(linha_dados)
        else:
            print(f"Linha ignorada por não ter colunas suficientes: {linha.text}")

    # Exibe os dados completos
    for dado in dados_completos:
        print(dado)

    return dados_completos


def salvar_dados_completos_planilha(dados_completos, nome_arquivo="teste.xlsx"):
    # Caminho para salvar o arquivo na área de trabalho
    caminho_arquivo = os.path.join(os.path.expanduser("~"), "Desktop", nome_arquivo)

    # Verifica se a planilha já existe
    if os.path.exists(caminho_arquivo):
        # Abre a planilha existente
        workbook = load_workbook(caminho_arquivo)
    else:
        # Cria uma nova planilha
        workbook = Workbook()

    # Seleciona a primeira folha da planilha
    folha = workbook.active
    folha.title = "Arrecadadores"

    # Adiciona o cabeçalho (apenas se a planilha for nova)
    if folha.max_row == 1:
        folha.append([
            "Arrecadador (Empresa)", "Qtde Títulos", "Operação",
            "Recolhimento CFEM", "% Recolhimento CFEM",  # Dados da segunda tabela
            "Ano", "Arrecadação por", "Ordenação por", "Substância Agrupadora", "Substância", "Região", "Estado"  # Dados gerais (primeira tabela)
        ])

    # Adiciona os dados capturados
    for dado in dados_completos:
        folha.append([
            dado["Arrecadador (Empresa)"],
            dado["Qtde Títulos"],
            dado["Operação"],
            dado["Recolhimento CFEM"],
            dado["% Recolhimento CFEM"],
            dado.get("Ano", ""),  # Dados gerais (adicionais)
            dado.get("Arrecadação por", ""),
            dado.get("Ordenação por", ""),
            dado.get("Substância Agrupadora", ""),
            dado.get("Substância", ""),
            dado.get("Região", ""),
            dado.get("Estado", "")
        ])

    # Salva a planilha
    workbook.save(caminho_arquivo)
    print(f"Dados salvos na planilha '{caminho_arquivo}' com sucesso!")
