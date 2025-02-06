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
from salvar_dados_planilha import capturar_todos_os_dados, salvar_dados_completos_planilha
from abrir_navegador import abrir_navegador
import numpy as np
import sys 

#Configurações
nome_arquivo = "Vai_dar_certo"
num_processos = 1  # 🔹 Instâncias paralelas dentro de cada grupo
subset_checkpoint_file = "subset_checkpoint.txt"  # 🔹 Checkpoint dos subsets

def carregar_checkpoint(arquivo):
    """Carrega os grupos ou subsets já processados do checkpoint."""
    if os.path.exists(arquivo):
        with open(arquivo, "r") as f:
            return set(f.read().splitlines())
    return set()
    
def salvar_checkpoint(arquivo, checkpoint_id):
    """Salva um grupo ou subset concluído no checkpoint."""
    print(f"✅ Salvando {checkpoint_id} no checkpoint...")
    with open(arquivo, "a") as f:
        f.write(f"{checkpoint_id}\n")  # Garante que o ID será salvo corretamente como string


def preencher_formulario(navegador, subs_agrupadora_valores):
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
                abrir_navegador()
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



def executar_robo(subs_agrupadora):
    """Executa o robô para uma única subs_agrupadora."""
    navegador = abrir_navegador()
    subs_agrupadora = selecionar_subs_agrupadora(func_navegador=navegador)
    subs_agrupadora_valores = [option.get_attribute("value") for option in subs_agrupadora.options if option.get_attribute("value")]
    preencher_formulario(navegador, subs_agrupadora_valores)
    navegador.quit()
    return subs_agrupadora_valores


if __name__ == "__main__":
    navegador = abrir_navegador()
    sleep(5)
    subs_agrupadora = selecionar_subs_agrupadora(func_navegador=navegador)
    subs_agrupadora_valores = [option.get_attribute("value") for option in subs_agrupadora.options if option.get_attribute("value")]
    
    print(f"✅ Subs Agrupadora Valores Encontrados: {subs_agrupadora_valores}")
    navegador.quit()

    checkpoint = carregar_checkpoint(subset_checkpoint_file)
    subs_pendentes = [subs_agrupada for subs_agrupada in subs_agrupadora_valores if subs_agrupada not in checkpoint]
    
    processos = []
    for subs_agrupadora in subs_pendentes:
        p = Process(target=executar_robo, args=(subs_agrupadora,))
        processos.append(p)
        p.start()

        if len(processos) >= num_processos:  # Se atingiu o limite, espera os processos terminarem antes de iniciar novos
            for p in processos:
                p.join()
            processos = []  # Esvazia a lista de processos concluídos

    # Garante que todos os processos terminem
    for p in processos:
        p.join()
        
    
    # mesclar_planilhas( f"{nome_arquivo}.xlsx")

                                    
                                
                                
      


                
                    
        
        
        