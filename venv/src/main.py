import pandas as pd
import matplotlib.pyplot as plt
from openpyxl import load_workbook
import re
from dateutil import parser
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
import dotenv
import os

# Carregar variáveis de ambiente
dotenv.load_dotenv(dotenv.find_dotenv())

# Carregar o arquivo Excel
excel = load_workbook('venv/src/Relacao_Produtos_e_Clientes_2024.xlsx')
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

def enviar_email(attachment_paths):
    smtp_server = 'smtp.gmail.com'
    smtp_port = 587  # Porta para TLS

    # Dados de autenticação
    remetente = os.getenv('remetente')
    senha_remetente = os.getenv('senha_remetente')
    destinatario = os.getenv('destinatario')

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

def verificar_numero(valor):
    try:
        float(valor[1:])
        return True
    except (ValueError, TypeError):
        return False

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

regex_data = r'\b\d{1,4}[-/]\d{1,2}[-/]\d{1,4}\b'

def data_convertida(data, regex_data):
    try:
        data_certa = re.search(regex_data, str(data))
        if data_certa:
            parsed_date = parser.parse(data_certa.group(), dayfirst=True)
            return parsed_date.strftime('%d/%m/%Y')
        else:
            return None
    except ValueError:
        return None

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

    lista_data.append(data_convertida(data, regex_data))
    lista_produtos.append(produto)

    if isinstance(valor, str) and verificar_numero(valor):
        valor_novo = re.sub(r'[$]', '', valor)  # Remover símbolo $
        lista_valor.append(float(valor_novo))
    elif isinstance(valor, (int, float)):
        lista_valor.append(valor)
    else:
        lista_valor.append(None)

    lista_regiao.append(regiao)
    lista_equipe.append(equipe)
    lista_cliente.append(cliente)
    lista_metPagamento.append(verificar_pagamento(metPagamento))

    if isinstance(desconto, (int, float)) and desconto >= 0:
        lista_desconto.append(desconto)
    else:
        lista_desconto.append(None)

# Criar DataFrame
df = pd.DataFrame({
    'Data da Venda': lista_data,
    'Produto': lista_produtos,
    'Valor da Venda': lista_valor,
    'Região': lista_regiao,
    'Equipe de Venda': lista_equipe,
    'Cliente': lista_cliente,
    'Método de Pagamento': lista_metPagamento,
    'Desconto': lista_desconto
})

def gerar_grafico(df, group_by, value_col, agg_func, kind, title, xlabel, ylabel, filename, legend_title=None, rotation=0):
    df_grouped = df.groupby(group_by)[value_col].agg(agg_func)
    
    if kind == 'line':
        df_grouped = df_grouped.fillna(0)  # Preencher NaN com 0 para gráfico de linha
    
    if isinstance(df_grouped.index, pd.MultiIndex):
        df_grouped = df_grouped.unstack()
    
    if kind == 'line':
        plt.figure(figsize=(12, 6))
        plt.plot(df_grouped.index, df_grouped.values, linestyle='-', color='b', linewidth=1, markersize=8)
    else:
        df_grouped.plot(kind=kind, figsize=(12, 6), width=0.8)

    plt.xlabel(xlabel)
    plt.ylabel(ylabel)
    plt.title(title)
    plt.xticks(rotation=rotation)
    if legend_title and kind != 'line':
        plt.legend(title=legend_title, bbox_to_anchor=(1.05, 1), loc='upper left')
    plt.tight_layout()
    plt.savefig(filename)
    plt.close()

# Gerar gráficos
gerar_grafico(df,
              group_by=['Produto', 'Região'],
              value_col='Valor da Venda',
              agg_func='mean',
              kind='bar',
              title='Valor Médio das Vendas por Produto e Região',
              xlabel='Produto',
              ylabel='Valor da Venda (Média)',
              filename='valor_medio_por_produto_e_regiao.png',
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
              filename='valor_medio_por_produto_e_equipe.png',
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
              filename='valor_total_por_regiao_e_metodo_pagamento.png',
              legend_title='Região',
              rotation=30)

gerar_grafico(df,
              group_by=['Equipe de Venda'],
              value_col='Desconto',
              agg_func='mean',
              kind='bar',
              title='Valor Médio dos Descontos Aplicados por Equipe',
              xlabel='Equipe',
              ylabel='Desconto Médio',
              filename='valor_medio_desconto_por_equipe.png',
              rotation=30)

gerar_grafico(df,
              group_by='Data da Venda',
              value_col='Valor da Venda',
              agg_func='mean',
              kind='line',
              title='Valor Médio das Vendas por Data da Venda',
              xlabel='Data da Venda',
              ylabel='Valor da Venda (Média)',
              filename='valor_medio_por_data.png',
              rotation=30)

# Enviar e-mail com os gráficos como anexo
enviar_email(['valor_medio_por_produto_e_regiao.png', 'valor_medio_por_produto_e_equipe.png', 'valor_total_por_regiao_e_metodo_pagamento.png', 'valor_medio_desconto_por_equipe.png', 'valor_medio_por_data.png'])