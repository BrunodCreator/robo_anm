from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from openpyxl import Workbook, load_workbook
import os
from datetime import datetime

# import preencher_formulario
from time import sleep


data_hora_atual = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
nome_arquivo = f"Maiores_Arrecadadores_{data_hora_atual}.xlsx"

   
    
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


def salvar_dados_completos_planilha(dados_completos, nome_arquivo=".xlsx"):
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
