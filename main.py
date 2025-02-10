from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from time import sleep
from multiprocessing import Process
import os
from datetime import datetime
from func.selecionar_combobox import selecionar_subs_agrupadora, selecionar_substancia, selecionar_regiao, selecionar_estado, selecionar_municipio
from func.selecionar_radio import selecionar_radio_empresa, selecionar_radio_operacao
from func.clicar_gera import clicar_gera
from func.salvar_dados_planilha import capturar_todos_os_dados, salvar_dados_completos_planilha
from func.abrir_navegador import abrir_navegador
from func.manipular_checkpoint import carregar_checkpoint, salvar_checkpoint
from func.preencher_formulario import preencher_formulario

if __name__ == "__main__":
    subset_checkpoint_file = "subset_checkpoint.txt"  # 🔹 Checkpoint dos subsets
    
    navegador = abrir_navegador()
    sleep(5)
    subs_agrupadora = selecionar_subs_agrupadora(func_navegador=navegador)
    subs_agrupadora_valores = [option.get_attribute("value") for option in subs_agrupadora.options if option.get_attribute("value")]
    
    print(f"✅ Subs Agrupadora Valores Encontrados: {subs_agrupadora_valores}")


    #Carrega checkpoint e filtra os pendentes
    checkpoint = carregar_checkpoint(subset_checkpoint_file)
    subs_pendentes = [subs_agrupada for subs_agrupada in subs_agrupadora_valores if subs_agrupada not in checkpoint]
    
    # 🔹 3️⃣ Processa cada subs_agrupadora uma por uma
    for subs_agrupadora in subs_pendentes:
        preencher_formulario(navegador, subs_pendentes,"subset_checkpoint.txt") 
        
    navegador.quit()
                      