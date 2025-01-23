from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from time import sleep
import os


def selecionar_radio_empresa(func_navegador):
    campo_radio = WebDriverWait(func_navegador, 20).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, '#ctl00_ContentPlaceHolder1_rdComparacao_5'))
    )

    # Clique no campo de rádio
    campo_radio.click()

    
def selecionar_radio_operacao(func_navegador):
    # Localize o campo de rádio usando o seletor CSS
    campo_radio_ordenacao = WebDriverWait(func_navegador, 20).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, '#ctl00_ContentPlaceHolder1_rdOrdenacao_0'))
    )

    # Clique no campo de rádio
    campo_radio_ordenacao.click()