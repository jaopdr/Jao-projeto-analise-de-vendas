from openpyxl import load_workbook
from datetime import datetime

excel = load_workbook('Relacao_Produtos_e_Clientes_2024.xlsx')

planilha = excel.active

def contar_produtos(file_path):
    excel = load_workbook(file_path)
    planilha = excel.active
    
    produtosA_contagem = {}
    
    for row in planilha.iter_rows(min_row=2, min_col=2, max_col=2, values_only=True):
        produto = row[0]
        
        if produto not in produtosA_contagem:
            produtosA_contagem[produto] = 0
        
        produtosA_contagem[produto] += 1
    
    return produtosA_contagem