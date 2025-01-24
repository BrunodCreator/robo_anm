from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from time import sleep
import os
# from func.selecionar_combobox import selecionar_subs_agrupadora, selecionar_substancia, selecionar_regiao, selecionar_estado, selecionar_municipio
# from func.selecionar_radio import selecionar_radio_empresa, selecionar_radio_operacao
# from func.clicar_gera import clicar_gera
from func.salvar_dados_planilha import capturar_todos_os_dados, salvar_dados_completos_planilha

#configurando o webdriver

service = Service(ChromeDriverManager().install())
navegador = webdriver.Chrome(service=service)


url = 'https://sistemas.anm.gov.br/arrecadacao/extra/relatorios/cfem/maiores_arrecadadores.aspx'
navegador.get(url)

sleep(15)

# Captura os dados combinados
dados_completos = capturar_todos_os_dados(func_navegador=navegador)

print(dados_completos)
# Salva os dados na planilha
salvar_dados_completos_planilha(dados_completos)