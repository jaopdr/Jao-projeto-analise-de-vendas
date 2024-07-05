import pandas as pd
import matplotlib.pyplot as plt
from openpyxl import load_workbook
import re
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage

# Carregar o arquivo Excel
excel = load_workbook('venv\scr\Relacao_Produtos_e_Clientes_2024.xlsx')
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

# Remover linhas onde Valor da Venda ou Desconto é None
df = df.dropna(subset=['Valor da Venda', 'Desconto'])

# Agrupar por Produto e Região e calcular a média do Valor da Venda
df_grupo1 = df.groupby(['Produto', 'Região'])['Valor da Venda'].mean().unstack()

# Plotar o gráfico de barras agrupadas por região
df_grupo1.plot(kind='bar', figsize=(12, 6), width=(0.8))
plt.xlabel('Produto')
plt.ylabel('Valor da Venda (Média)')
plt.title('Valor Médio das Vendas por Produto e Região')
plt.xticks(rotation=30)
plt.legend(title='Região', bbox_to_anchor=(1.05, 1), loc='upper left')
plt.tight_layout()
plt.savefig('plot1.png')

# Agrupar por Produto e Equipe de Venda e calcular a média do Valor da Venda
df_grupo2 = df.groupby(['Produto', 'Equipe de Venda'])['Valor da Venda'].mean().unstack()

# Plotar o gráfico de barras agrupadas por equipe
df_grupo2.plot(kind='bar', figsize=(12, 6), width=(0.8))
plt.xlabel('Produto')
plt.ylabel('Valor da Venda (Média)')
plt.title('Valor Médio das Vendas por Produto e Equipe')
plt.xticks(rotation=30)
plt.legend(title='Equipe', bbox_to_anchor=(1.05, 1), loc='upper left')
plt.tight_layout()
plt.savefig('plot2.png')

# Agrupar por Região e Método de Pagamento e calcular a soma do Valor da Venda
df_metodo_pagamento = df.groupby(['Método de Pagamento', 'Região'])['Valor da Venda'].mean().unstack()

# Plotar o gráfico de barras do valor total de vendas por método de pagamento e região
df_metodo_pagamento.plot(kind='bar', figsize=(12, 6), width=(0.8))
plt.xlabel('Método de Pagamento')
plt.ylabel('Valor Total de Vendas')
plt.title('Valor Total de Vendas por Região e Método de Pagamento')
plt.xticks(rotation=45)
plt.legend(title='Região', bbox_to_anchor=(1.05, 1), loc='upper left')
plt.tight_layout()
plt.savefig('plot3.png')

# Agrupar por Equipe de Venda e calcular a média do Desconto
df_equipe_desconto = df.groupby(['Equipe de Venda'])['Desconto'].mean()

# Definir cores para cada equipe
colors = ['blue', 'red', 'orange', 'green', 'purple'] * (len(df_equipe_desconto) // 5 + 1)
colors = colors[:len(df_equipe_desconto)]

# Plotar o gráfico de barras da média de desconto por equipe
plt.figure(figsize=(12, 6))
df_equipe_desconto.plot(kind='bar', width=0.8, color=colors)
plt.xlabel('Equipe de Venda')
plt.ylabel('Desconto (Média)')
plt.title('Média de Desconto por Equipe de Venda')
plt.xticks(rotation=45)
plt.tight_layout()
plt.savefig('plot4.png')

# Enviar e-mail com os gráficos como anexo
enviar_email(['plot1.png', 'plot2.png', 'plot3.png', 'plot4.png'])