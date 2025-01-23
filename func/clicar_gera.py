from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

def clicar_gera(func_navegador):
    botao_gera = WebDriverWait(func_navegador, 20).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, '#ctl00_ContentPlaceHolder1_btnGera'))
    )
    botao_gera.click()
    print("Botão de geração clicado com sucesso!")

    # Aguardar que os resultados da pesquisa sejam carregados
    WebDriverWait(func_navegador, 30).until(
        EC.presence_of_all_elements_located((By.CSS_SELECTOR, '#ctl00_ContentPlaceHolder1_dvResultado > table.tabelaRelatorio'))  # Ajuste o seletor aqui
    )
    print("Resultados carregados com sucesso!")