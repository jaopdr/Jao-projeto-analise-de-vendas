from openpyxl import load_workbook

excel = load_workbook('Relacao_Produtos_e_Clientes_2024.xlsx')

planilha = excel.active

def contar_produtos():
    produtosA_contagem = 0
    produtosB_contagem = 0
    produtosC_contagem = 0
    produtosX_contagem = 0
    produtosY_contagem = 0

    for row in planilha.iter_rows(min_row=2, min_col=2, max_col=2, values_only=True):
        produto = row[0]

        if produto == "Produto A":
            produtosA_contagem += 1
        elif produto == "Produto B":
            produtosB_contagem += 1
        elif produto == "Produto C":
            produtosC_contagem += 1
        elif produto == "Produto X":
            produtosX_contagem += 1
        elif produto == "Produto Y":
            produtosY_contagem += 1

    return {
        "Produto A": produtosA_contagem,
        "Produto B": produtosB_contagem,
        "Produto C": produtosC_contagem,
        "Produto X": produtosX_contagem,
        "Produto Y": produtosY_contagem
    }

contagem_produtos = contar_produtos()
print(contagem_produtos)
