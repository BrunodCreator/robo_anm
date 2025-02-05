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
import numpy as np
import sys 

#Configurações
nome_arquivo = "Teste_seis"
num_grupos = 12   # 🔹 Divisão principal da lista
num_processos = 6  # 🔹 Instâncias paralelas dentro de cada grupo
checkpoint_file = "checkpoint.txt" #Arquivo de checkpoint
subset_checkpoint_file = "subset_checkpoint.txt"  # 🔹 Checkpoint dos subsets

def carregar_checkpoint(arquivo):
    """Carrega os grupos ou subsets já processados do checkpoint."""
    if os.path.exists(arquivo):
        with open(arquivo, "r") as f:
            return set(f.read().splitlines())
    return set()
    
def salvar_checkpoint(arquivo, grupo_id):
    """Salva um grupo ou subset concluído no checkpoint"""
    print(f'Salvando {id} no checkpoint...')
    with open(arquivo, "a") as f:
        f.write(f'{id}\n')

def preencher_formulario(navegador, subs_agrupadora_sublist, grupo_id, subset_id):
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
            try:
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
            except Exception as e:
                print(f'Não consegui recuperar do erro usando refresh, abrindo navegador novamente...')
                navegador.quit()
                abrir_navegador()
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
        # Salvar o progresso após cada substância
        salvar_checkpoint(subset_checkpoint_file, f"{grupo_id}_{subset_id}_{substancia_agrupadora}")



def executar_robo(subset, grupo_id, subset_id):
    """Executa o robô para um subconjunto de substâncias"""
    navegador = abrir_navegador()
    preencher_formulario(navegador, subset, grupo_id, subset_id)
    navegador.quit()


if __name__ == "__main__":
    navegador = abrir_navegador()
    sleep(5)
    subs_agrupadora = selecionar_subs_agrupadora(func_navegador=navegador)
    subs_agrupadora_valores = [option.get_attribute("value") for option in subs_agrupadora.options if option.get_attribute("value")]
    navegador.quit()

    grupos = np.array_split(subs_agrupadora_valores, num_grupos)
    grupos = [grupo.tolist() for grupo in grupos]
    checkpoint = carregar_checkpoint(checkpoint_file)
    subset_checkpoint = carregar_checkpoint(subset_checkpoint_file)

    for grupo_id, grupo in enumerate(grupos):
        if str(grupo_id) in checkpoint:
            print(f"✅ Grupo {grupo_id} já processado. Pulando...")
            continue

        print(f"🚀 Iniciando processamento do Grupo {grupo_id + 1}/{num_grupos}...")
        subsets = np.array_split(grupo, num_processos)
        subsets = [sub.tolist() for sub in subsets]

        processos = []
        for subset_id, subset in enumerate(subsets):
            subset_key = f"{grupo_id}_{subset_id}"
            if subset_key in subset_checkpoint:
                print(f"✅ Subset {subset_id} do Grupo {grupo_id} já processado. Pulando...")
                continue
            p = Process(target=executar_robo, args=(subset, grupo_id, subset_id))
            processos.append(p)
            p.start()

        for p in processos:
            p.join()

        salvar_checkpoint(checkpoint_file, grupo_id)

    arquivos_para_mesclar = [f"{nome_arquivo}_Grupo{i}_Subset{j}.xlsx" for i in range(num_grupos) for j in range(num_processos) if f"{i}_{j}" in subset_checkpoint]
    mesclar_planilhas(arquivos_para_mesclar, nome_arquivo_final=f"{nome_arquivo}.xlsx")

                                    
                                
                                
      


                
                    
        
        
        