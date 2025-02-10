from func.clicar_gera import clicar_gera
from func.salvar_dados_planilha import capturar_todos_os_dados
from func.selecionar_combobox import selecionar_estado
from func.selecionar_combobox import selecionar_municipio

def selecionar_todos_os_estados(navegador):
    """Seleciona a opção 'Todos os Estados' usando a função selecionar_estado"""
    estados = selecionar_estado(navegador)
    estados.select_by_value("Todos os Estados")
    print("🗂️ Selecionado: Todos os Estados")



def selecionar_todos_os_municipios(navegador):
    """Seleciona a opção 'Todas os Município' usando a função selecionar_municipio."""
    municipios = selecionar_municipio(navegador)
    municipios.select_by_value("Todas os Município")
    print("🗂️ Selecionado: Todas os Município")



def verificar_existencia_de_dados(navegador, subs):
    """Verifica se há dados após clicar em 'Gerar'. Retorna True se houver dados, False caso contrário."""
    try:
        selecionar_todos_os_estados(navegador)  # Seleciona "Todos os Estados"
        
        clicar_gera(navegador)  # Clica no botão "Gerar"
        
        dados_completos = capturar_todos_os_dados(func_navegador=navegador)  # Captura os dados
        
        if dados_completos:
            print(f'✅ Dados encontrados para a substancia {subs}.')
            return True
        else:
            print('❌ Nenhum dado encontrado.')
            return False
    except Exception as e:
        print(f"⚠️ Erro ao verificar existência de dados: {e}")
        return False



def verificar_existencia_de_dados_por_estado(navegador,estado):
    """Verifica se há dados em cada estado após clicar em 'Gerar'. Retorna True se houver dados, False caso contrário."""
    try:
        selecionar_todos_os_municipios(navegador)
        
        clicar_gera(navegador)
        
        dados_completos = capturar_todos_os_dados(func_navegador=navegador)  # Captura os dados
        
        if dados_completos:
            print(f'✅ Dados de municipio ({estado}) encontrados.')
            return True
        else:
            print(f'❌ Nenhum dado encontrado para o município {estado}.')
            return False
    except Exception as e:
        print(f"⚠️ Erro ao verificar existência de dados: {e}")
        return False
        