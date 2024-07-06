import pandas as pd
import matplotlib.pyplot as plt
from openpyxl import load_workbook
import re
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage

# Carregar o arquivo Excel
excel = load_workbook('venv/src/Relacao_Produtos_e_Clientes_2024.xlsx')
planilha = excel.active

# Listas para armazenar os dados
lista_produtos = []
lista_valor = []
lista_regiao = []
lista_equipe = []
lista_cliente = []
lista_metPagamento = []
lista_desconto = []

def enviar_email(attachment_paths):
    # Configurações do servidor SMTP
    smtp_server = 'smtp.gmail.com'
    smtp_port = 587  # Porta para TLS
    
    # Dados de autenticação
    remetente = 'jpedro.seze@gmail.com'
    senha_remetente = 'spavzcyfkwphhrms'
    destinatario = 'joaosezerino.dev@gmail.com'

    # Construir o e-mail
    msg = MIMEMultipart()
    msg['From'] = remetente
    msg['To'] = destinatario
    msg['Subject'] = 'Gráficos Anexos'

    # Anexar imagens ao e-mail
    for path in attachment_paths:
        with open(path, 'rb') as attachment:
            image = MIMEImage(attachment.read())
            image.add_header('Content-Disposition', f'attachment; filename={path}')
            msg.attach(image)

    # Enviar e-mail usando SMTP
    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(remetente, senha_remetente)
        server.sendmail(remetente, destinatario, msg.as_string())
        print('E-mail enviado com sucesso!')
    except Exception as e:
        print(f'Falha ao enviar o e-mail: {e}')
    finally:
        server.quit()

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

# Função para gerar gráficos
def gerar_grafico(df, group_by, value_col, agg_func, kind, title, xlabel, ylabel, filename, legend_title=None, rotation=0):
    df_grouped = df.groupby(group_by)[value_col].agg(agg_func)
    if isinstance(df_grouped.index, pd.MultiIndex):
        df_grouped = df_grouped.unstack()
    df_grouped.plot(kind=kind, figsize=(12, 6), width=0.8)
    plt.xlabel(xlabel)
    plt.ylabel(ylabel)
    plt.title(title)
    plt.xticks(rotation=rotation)
    if legend_title:
        plt.legend(title=legend_title, bbox_to_anchor=(1.05, 1), loc='upper left')
    plt.tight_layout()
    plt.savefig(filename)
    plt.close()  # Fechar a figura atual para evitar sobreposição

# Iterar sobre as linhas da planilha e preencher as listas
for row in planilha.iter_rows(min_row=2, values_only=True):
    produto = row[1]
    valor = row[2]
    regiao = row[3]
    equipe = row[4]
    cliente = row[5]
    metPagamento = row[6]
    desconto = row[7]

    # Adicionar produto à lista de produtos
    lista_produtos.append(produto)

    # Verificar se o valor é numérico e adicionar à lista de valores
    if verificar_numero(valor):
        valor_novo = re.sub(r'[$]', '', valor)  # Remover símbolo $
        lista_valor.append(float(valor_novo))
    else:
        lista_valor.append(None)

    # Adicionar região à lista de regiões
    lista_regiao.append(regiao)

    # Adicionar equipe à lista de equipes
    lista_equipe.append(equipe)

    # Adicionar cliente à lista de clientes
    lista_cliente.append(cliente)

    # Substituir o método de pagamento e adicionar à lista método pagamento
    metodo_pagamento_corrigido = verificar_pagamento(metPagamento)
    lista_metPagamento.append(metodo_pagamento_corrigido)

    # Verificar se o desconto é positivo e adicionar à lista de descontos
    if desconto >= 0:
        lista_desconto.append(desconto)
    else:
        lista_desconto.append(None)

# Criar DataFrame
df = pd.DataFrame({
    'Produto': lista_produtos,
    'Valor da Venda': lista_valor,
    'Região': lista_regiao,
    'Equipe de Venda': lista_equipe,
    'Cliente': lista_cliente,
    'Método de Pagamento': lista_metPagamento,
    'Desconto': lista_desconto
})

# Gerar gráficos
gerar_grafico(df, 
              group_by=['Produto', 'Região'], 
              value_col='Valor da Venda', 
              agg_func='mean', 
              kind='bar', 
              title='Valor Médio das Vendas por Produto e Região', 
              xlabel='Produto', 
              ylabel='Valor da Venda (Média)', 
              filename='plot1.png', 
              legend_title='Região', 
              rotation=30)

gerar_grafico(df, 
              group_by=['Produto', 'Equipe de Venda'], 
              value_col='Valor da Venda', 
              agg_func='mean', 
              kind='bar', 
              title='Valor Médio das Vendas por Produto e Equipe', 
              xlabel='Produto', 
              ylabel='Valor da Venda (Média)', 
              filename='plot2.png', 
              legend_title='Equipe', 
              rotation=30)

gerar_grafico(df, 
              group_by=['Método de Pagamento', 'Região'], 
              value_col='Valor da Venda', 
              agg_func='sum', 
              kind='bar', 
              title='Valor Total de Vendas por Região e Método de Pagamento', 
              xlabel='Método de Pagamento', 
              ylabel='Valor Total de Vendas', 
              filename='plot3.png', 
              legend_title='Região', 
              rotation=45)

gerar_grafico(df,
              group_by=['Equipe de Venda'],
              value_col='Desconto',
              agg_func='mean',
              kind='bar',
              title='Valor Médio dos Descontos Aplicados por Equipe',
              xlabel='Equipe',
              ylabel='Desconto Médio',
              filename='plot4.png',
              rotation=30)

# # Enviar e-mail com os gráficos como anexo
# enviar_email(['plot1.png', 'plot2.png', 'plot3.png', 'plot4.png'])