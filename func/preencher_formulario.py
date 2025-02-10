from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from time import sleep
import os
from datetime import datetime
from func.selecionar_combobox import selecionar_subs_agrupadora, selecionar_substancia, selecionar_regiao, selecionar_estado, selecionar_municipio
from func.selecionar_radio import selecionar_radio_empresa, selecionar_radio_operacao
from func.clicar_gera import clicar_gera
from func.salvar_dados_planilha import capturar_todos_os_dados, salvar_dados_completos_planilha
from func.abrir_navegador import abrir_navegador
from func.manipular_checkpoint import carregar_checkpoint, salvar_checkpoint
from func.verificar_existencia_de_dados import verificar_existencia_de_dados, verificar_existencia_de_dados_por_estado



nome_arquivo = "Vai_dar_certo"

def preencher_formulario(navegador, subs_agrupadora_valores, subset_checkpoint_file):
    """Preenche o formulário apenas para um subconjunto de dados"""
# Dividir os dados de `subs_agrupadora` em subconjuntos
    checkpoint = carregar_checkpoint(subset_checkpoint_file)
    
    for subs_agrupada in subs_agrupadora_valores:
        if subs_agrupada in checkpoint:
            print(f"🔄 {subs_agrupada} já processado. Pulando...")
            continue  # Pula se já estiver no checkpoint
        try:
            if subs_agrupada == 'Todas as Agrupadoras':
                continue    
            subs_agrupadora = selecionar_subs_agrupadora(func_navegador=navegador)
            
            #LINHA QUE SELECIONA O CAMPO NO COMBOBOX
            subs_agrupadora.select_by_value(subs_agrupada)
            print(f'subs_agrupadora: {subs_agrupada}')
            
            substancia = selecionar_substancia(func_navegador=navegador)
            substancia_valores = [option.get_attribute("value") for option in substancia.options if option.get_attribute("value")]
        except Exception as e:
            try:
                print(f'Erro ao selecionar a subs_agrupadora erro: {e}')
                navegador.refresh()
                if subs == 'Todas as Agrupadoras':
                    continue    
                subs_agrupadora = selecionar_subs_agrupadora(func_navegador=navegador)
                #LINHA QUE SELECIONA O CAMPO NO COMBOBOX
                subs_agrupadora.select_by_value(subs_agrupada)
                print(f'subs_agrupadora: {subs_agrupada}')
                
                substancia = selecionar_substancia(func_navegador=navegador)
                substancia_valores = [option.get_attribute("value") for option in substancia.options if option.get_attribute("value")]
            except Exception as e:
                print(f'Não consegui recuperar do erro usando refresh, abrindo navegador novamente...')
                navegador.quit()
                navegador = abrir_navegador()
                if subs_agrupada == 'Todas as Agrupadoras':
                    continue    
                subs_agrupadora = selecionar_subs_agrupadora(func_navegador=navegador)
                #LINHA QUE SELECIONA O CAMPO NO COMBOBOX
                subs_agrupadora.select_by_value(subs_agrupada)
                print(f'subs_agrupadora: {subs_agrupada}')
                
                substancia = selecionar_substancia(func_navegador=navegador)
                substancia_valores = [option.get_attribute("value") for option in substancia.options if option.get_attribute("value")]
                
        for subs in substancia_valores:
            try:
                if subs == 'Todas as Substância':
                    continue
                substancia = selecionar_substancia(func_navegador=navegador)
                #LINHA QUE SELECIONA O CAMPO NO COMBOBOX
                print(f'Os valores disponíveis de substancias para a subs.agrupadora:{subs} são :{substancia_valores}')
                substancia.select_by_value(subs)
                print(f'substancia_interna: {subs}')
                
                #SELECIONAR A REGIÃO CENTRO-OESTE
                selecionar_regiao(func_navegador=navegador)
                
                if not verificar_existencia_de_dados(navegador, subs=subs):
                    print(f'❌ Sem dados para a substância {subs}. Pulando para a próxima.')
                    continue #Pula para a próxima substância
                
                estados = selecionar_estado(func_navegador=navegador)
                estados_valores = [option.get_attribute("value") for option in estados.options if option.get_attribute("value")]
                    
            except Exception as e:
                print(f'Erro ao selecionar a substancia erro: {e}')
                navegador.refresh()
                if subs == 'Todas as Substância':
                    continue
                substancia = selecionar_substancia(func_navegador=navegador)
                #LINHA QUE SELECIONA O CAMPO NO COMBOBOX
                print(f'Os valores disponíveis de substancias para a subs.agrupadora:{subs} são :{substancia_valores}')
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
                    
                    if not verificar_existencia_de_dados_por_estado(navegador, estado=estado):
                        print(f'❌ Sem dados para a substância {subs}. Pulando para a próxima.')
                        continue #Pula para a próxima substância
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
                        print(dados_completos)
                        print(datetime.now().strftime("%H-%M-%S"))
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
                        print(dados_completos)
                        print(datetime.now().strftime("%H-%M-%S"))
                        salvar_dados_completos_planilha(dados_completos, nome_arquivo=f"{nome_arquivo}.xlsx")
        
        # Somente salva no checkpoint após concluir todas as substâncias
        salvar_checkpoint(subset_checkpoint_file, subs_agrupada)






                                
                                
      


                
                    
        
        
        