from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options

#configurando o webdriver
def abrir_navegador():
    try:
        chrome_options = Options()
        # Ignorar erros de SSL
        chrome_options.add_argument('--ignore-certificate-errors')
        chrome_options.add_argument('--allow-insecure-localhost')

        print("Iniciando o navegador...")
        service = Service(ChromeDriverManager().install())
        navegador = webdriver.Chrome(service=service, options=chrome_options)
        url = 'https://sistemas.anm.gov.br/arrecadacao/extra/Relatorios/cfem/arrecadadores.aspx'
        navegador.get(url)
        print("Navegador iniciado com sucesso!")
        return navegador
    except Exception as e:
        print(f"Erro ao iniciar o navegador: {e}")
        return None
