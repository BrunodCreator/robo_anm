from openpyxl import workbook, load_workbook
import os
from datetime import datetime

def capturar_dados_table():
    ano = datetime.now().year

    
    
    
    
    
    
    
    
    
    
    
    
caminho_planilha = os.path.join(os.path.expanduser("~"), "Desktop", "Arrecadadores CFEM.xlsx")

workbook = load_workbook(caminho_planilha)
sheet = workbook["Folha1"]


sheet['A1'] = 'Ano'
sheet['B1'] = 'Subs.Agrupadora'
sheet['C1'] = 'Substância'
sheet['D1'] = 'Região'
sheet['E1'] = 'Estado'
sheet['F1'] = 'Municipio'
sheet['G1'] = 'Arrecadador (Empresa)'
sheet['H1'] = 'Qtde Títulos'
sheet['K1'] = 'Operação'
sheet['I1'] = 'Recolhimento CFEM'
sheet['J1'] = '% Recolhimento CFEM'

linha_excel = 2