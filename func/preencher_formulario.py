from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from multiprocessing import Process
from time import sleep
import os
from datetime import datetime
from selecionar_combobox import selecionar_subs_agrupadora, selecionar_substancia, selecionar_regiao, selecionar_estado, selecionar_municipio
from selecionar_radio import selecionar_radio_empresa, selecionar_radio_operacao
from clicar_gera import clicar_gera
from salvar_dados_planilha import capturar_todos_os_dados, salvar_dados_completos_planilha, mesclar_planilhas
from abrir_navegador import abrir_navegador

nome_arquivo = "Teste_seis"

# Definir o número de processos (quantas instâncias do robô rodarão ao mesmo tempo)
num_processos = 6


def preencher_formulario(navegador, subs_agrupadora_sublist):
    """Preenche o formulário apenas para um subconjunto de dados"""
# Dividir os dados de `subs_agrupadora` em subconjuntos

    for substancia_agrupadora in subs_agrupadora_sublist:
        try:
            if substancia_agrupadora == 'Todas as Agrupadoras':
                continue    
            subs_agrupadora = selecionar_subs_agrupadora(func_navegador=navegador)
            #LINHA QUE SELECIONA O CAMPO NO COMBOBOX
            subs_agrupadora.select_by_value(substancia_agrupadora)
            print(f'subs_agrupadora: {substancia_agrupadora}')
            
            substancia = selecionar_substancia(func_navegador=navegador)
            substancia_valores = [option.get_attribute("value") for option in substancia.options if option.get_attribute("value")]
        except Exception as e:
            print(f'Erro ao selecionar a subs_agrupadora erro: {e}')
            navegador.refresh()
            if substancia_agrupadora == 'Todas as Agrupadoras':
                continue    
            subs_agrupadora = selecionar_subs_agrupadora(func_navegador=navegador)
            #LINHA QUE SELECIONA O CAMPO NO COMBOBOX
            subs_agrupadora.select_by_value(substancia_agrupadora)
            print(f'subs_agrupadora: {substancia_agrupadora}')
            
            substancia = selecionar_substancia(func_navegador=navegador)
            substancia_valores = [option.get_attribute("value") for option in substancia.options if option.get_attribute("value")]
        for subs in substancia_valores:
            try:
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
            except Exception as e:
                print(f'Erro ao selecionar a substancia erro: {e}')
                navegador.refresh()
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
                try:
                    if estado == 'Todos os Estados':
                        continue
                    estados = selecionar_estado(func_navegador=navegador)
                    estados.select_by_value(estado)
                    print(f'Estado: {estado}')
                    
                    municipios = selecionar_municipio(func_navegador=navegador)
                    municipios_valores = [option.get_attribute("value") for option in municipios.options if option.get_attribute("value")]
                except Exception as e:
                    print(f'Erro ao selecionar o estado erro: {e}')
                    navegador.refresh()
                    if estado == 'Todos os Estados':
                        continue
                    estados = selecionar_estado(func_navegador=navegador)
                    estados.select_by_value(estado)
                    print(f'Estado: {estado}')
                    
                    municipios = selecionar_municipio(func_navegador=navegador)
                    municipios_valores = [option.get_attribute("value") for option in municipios.options if option.get_attribute("value")]
                for municipio in municipios_valores:
                    try:
                        if municipio == 'Todas os Município':
                            continue
                        municipios = selecionar_municipio(func_navegador=navegador)
                        municipios.select_by_value(municipio)
                        print(f'Municipio: {municipio}')
                        
                        selecionar_radio_empresa(func_navegador=navegador)
                        
                        selecionar_radio_operacao(func_navegador=navegador)
                
                        clicar_gera(func_navegador=navegador)
                        
                         # Capturar e salvar os dados
                        dados_completos = capturar_todos_os_dados(func_navegador=navegador)
                        salvar_dados_completos_planilha(dados_completos, nome_arquivo=f"{nome_arquivo}.xlsx")
                    except Exception as e:
                        print(f'Erro ao selecionar os municípios erro: {e} ')
                        navegador.refresh()
                        if municipio == 'Todas os Município':
                            continue
                        municipios = selecionar_municipio(func_navegador=navegador)
                        municipios.select_by_value(municipio)
                        print(f'Municipio: {municipio}')
                        
                        selecionar_radio_empresa(func_navegador=navegador)
                        
                        selecionar_radio_operacao(func_navegador=navegador)
                
                        clicar_gera(func_navegador=navegador)

                         # Capturar e salvar os dados
                        dados_completos = capturar_todos_os_dados(func_navegador=navegador)
                        salvar_dados_completos_planilha(dados_completos, nome_arquivo=f"{nome_arquivo}.xlsx")



def executar_robo(subset):
    """Executa o robô para um subconjunto de substâncias"""
    navegador = abrir_navegador()
    preencher_formulario(navegador, subset)
    navegador.quit()


if __name__ == '__main__':
    # Abrir o navegador temporário apenas para capturar os valores da `subs_agrupadora`
    navegador = abrir_navegador()
    sleep(5)

    # Capturar valores das `comboboxes`
    subs_agrupadora = selecionar_subs_agrupadora(func_navegador=navegador)
    subs_agrupadora_valores = [option.get_attribute("value") for option in subs_agrupadora.options if option.get_attribute("value")]

    navegador.quit()  # Fechar o navegador temporário

    # Dividir os dados de `subs_agrupadora` em subconjuntos
    subsets = [subs_agrupadora_valores[i::num_processos] for i in range(num_processos)]

    # Criar e iniciar os processos para cada conjunto 
    processos = []
    for subset in subsets:
        p = Process(target=executar_robo, args=(subset,))  # Corrigido para passar apenas um subconjunto
        processos.append(p)
        p.start()

    for p in processos:
        p.join()
        
mesclar_planilhas(f"{nome_arquivo}.xlsx")

                                    
                                
                                
      


                
                    
        
        
        