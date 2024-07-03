from openpyxl import load_workbook

# Carregar o arquivo Excel
excel = load_workbook('Relacao_Produtos_e_Clientes_2024.xlsx')
planilha = excel.active

# Listas para armazenar os dados
lista_data = []
lista_produtos = []
lista_valor = []
lista_regiao = []
lista_equipe = []
lista_cliente = []
lista_metPagamento = []
lista_desconto = []

# Função para verificar se o valor a partir da segunda letra é numérico
def verificar_numero(valor):
    try:
        float(valor[1:])
        return True
    except (ValueError, TypeError):
        return False
    
# Dicionário de substituição para métodos de pagamento
def verificar_pagamento(metPagamento):
    substituicao = {
        'cartao_credito': ['Cartão de Crédito', 'Cred.'], 
        'cartao_debito': ['Cartão de Débito'],
        'transf_bancaria': ['Transferência Bancária', 'Tran. Bancária'],
        'dinheiro': ['Dinheiro'],
        'cheque': ['Cheque']
    }
    for key, values in substituicao.items():
        if metPagamento in values:
            return key
    return metPagamento

def verificar_data(row):


# Iterar sobre as linhas da planilha e preencher as listas
for row in planilha.iter_rows(min_row=2, values_only=True):
    data = row[0]
    produto = row[1]
    valor = row[2]
    regiao = row[3]
    equipe = row[4]
    cliente = row[5]
    metPagamento = row[6]
    desconto = row[7]

    # Verificar e padronizar as datas e adicionar à lista de datas
    lista_data.append(data)

    # Adicionar produto à lista de produtos
    lista_produtos.append(produto)

    # Adicionar valor à lista de valores somente se passar na verificação
    if verificar_numero(valor):
        lista_valor.append(valor)

    # Adicionar regiao à lista de regioes
    lista_regiao.append(regiao)

    # Adicionar equipe à lista de equipes
    lista_equipe.append(equipe)

    # Adicionar cliente à lista de clientes
    lista_cliente.append(cliente)

    # Substituir o método de pagamento e adicionar à lista metodo pagamento
    metodo_pagamento_corrigido = verificar_pagamento(metPagamento)
    lista_metPagamento.append(metodo_pagamento_corrigido)

    # Verificar o se o valor do desconto é condizente e adicionar à lista desconto
    if desconto >= 0:
        lista_desconto.append(desconto)

# Imprimir as listas para verificar os resultados
print("Datas: ", lista_data)
print("Produtos:", lista_produtos)
print("Valores:", lista_valor)
print("Regiões:", lista_regiao)
print("Equipes:", lista_equipe)
print("Clientes:", lista_cliente)
print("Metódo Pagamento: ", lista_metPagamento)
print("Desconto: ", lista_desconto)