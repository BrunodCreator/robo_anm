from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from time import sleep
import os

#SELEÇÃO DO CAMPO SUBS.AGRUPADORA
def selecionar_subs_agrupadora(func_navegador):
    subs_agrupadora_dropdow = WebDriverWait(func_navegador, 20).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, '#ctl00_ContentPlaceHolder1_subs_agrupadora'))
            )
    subs_agrupadora = Select(subs_agrupadora_dropdow)
    
    return subs_agrupadora
    #FIM SELEÇÃO DO CAMPO SUBS.AGRUPADORA
    
    
    
def selecionar_substancia(func_navegador):
    #SELEÇÃO DO CAMPO SUBSTANCIA
    substancia_dropdow = WebDriverWait(func_navegador, 20).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, '#ctl00_ContentPlaceHolder1_Substancia'))
    )
    substancia = Select(substancia_dropdow)
    return substancia
    #FIM SELEÇÃO DO CAMPO SUBSTANCIA
    
    
    
def selecionar_regiao(func_navegador):
    regiao_dropdow = WebDriverWait(func_navegador, 20).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, '#ctl00_ContentPlaceHolder1_regiao'))
    )
    regiao = Select(regiao_dropdow)
    regiao.select_by_value("CO")
    
    
def selecionar_estado(func_navegador):
    estado_dropdow = WebDriverWait(func_navegador, 20).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, '#ctl00_ContentPlaceHolder1_Estado'))
    )
    estados = Select(estado_dropdow)
    return estados


def selecionar_municipio(func_navegador):
    municipio_dropdow = WebDriverWait(func_navegador, 20).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, '#ctl00_ContentPlaceHolder1_Municipio'))
    )
    municipios = Select(municipio_dropdow)
    return municipios



