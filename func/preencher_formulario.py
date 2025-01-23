from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from time import sleep
import os
from selecionar_combobox import selecionar_subs_agrupadora, selecionar_substancia, selecionar_regiao, selecionar_estado, selecionar_municipio
from selecionar_radio import selecionar_radio_empresa, selecionar_radio_operacao
from clicar_gera import clicar_gera
from salvar_dados_planilha import extrair_dados
#configurando o webdriver
service = Service(ChromeDriverManager().install())
navegador = webdriver.Chrome(service=service)


url = 'https://sistemas.anm.gov.br/arrecadacao/extra/relatorios/cfem/maiores_arrecadadores.aspx'
navegador.get(url)



subs_agrupadora = selecionar_subs_agrupadora(func_navegador=navegador)
subs_agrupadora_valores = [option.get_attribute("value") for option in subs_agrupadora.options if option.get_attribute("value")]

for substancia_agrupadora in subs_agrupadora_valores:
    if substancia_agrupadora == 'Todas as Agrupadoras':
        continue    
    subs_agrupadora = selecionar_subs_agrupadora(func_navegador=navegador)
    #LINHA QUE SELECIONA O CAMPO NO COMBOBOX
    subs_agrupadora.select_by_value(substancia_agrupadora)
    print(f'subs_agrupadora: {substancia_agrupadora}')
    
    substancia = selecionar_substancia(func_navegador=navegador)
    substancia_valores = [option.get_attribute("value") for option in substancia.options if option.get_attribute("value")]
    for subs in substancia_valores:
        if subs == 'Todas as Substância':
            continue
        substancia = selecionar_substancia(func_navegador=navegador)
        #LINHA QUE SELECIONA O CAMPO NO COMBOBOX
        print(f'Os valores disponíveis de substancias para a subs.agrupadora:{substancia_agrupadora} são :{substancia_valores}')
        substancia.select_by_value(subs)
        print(f'substancia_interna: {subs}')
         
        #SELECIONAR A REGIÃO CENTRO-OESTE
        selecionar_regiao(func_navegador=navegador)
        
        estados = selecionar_estado(func_navegador=navegador)
        estados_valores = [option.get_attribute("value") for option in estados.options if option.get_attribute("value")]
        for estado in estados_valores:
            if estado == 'Todos os Estados':
                continue
            estados = selecionar_estado(func_navegador=navegador)
            estados.select_by_value(estado)
            print(f'Estado: {estado}')
            
            municipios = selecionar_municipio(func_navegador=navegador)
            municipios_valores = [option.get_attribute("value") for option in municipios.options if option.get_attribute("value")]
            for municipio in municipios_valores:
                if municipio == 'Todas os Município':
                    continue
                municipios = selecionar_municipio(func_navegador=navegador)
                municipios.select_by_value(municipio)
                print(f'Municipio: {municipio}')
                
                selecionar_radio_empresa(func_navegador=navegador)
                print("Campo de empresa selecionado com sucesso!")
                
                selecionar_radio_operacao(func_navegador=navegador)
                print('Campo de operação selecionado com sucesso')
                
                clicar_gera(func_navegador=navegador)
            
                
        
        
        