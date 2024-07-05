# Análise de Vendas e Envio de Relatórios Automatizados

## Descrição do Projeto

Este projeto realiza a análise de dados de vendas a partir de um arquivo Excel, gera gráficos informativos sobre o desempenho de produtos, equipes de venda, métodos de pagamento e descontos médios por equipe. Após a geração dos gráficos, os mesmos são enviados por e-mail como anexos.

## Objetivo

Este programa visa eliminar processos repetitivos, reduzir erros e agilizar operações para aumentar a eficiência e a precisão. Ao automatizar tarefas, ele busca melhorar a gestão, economizar tempo e permitir o direcionamento de recursos para áreas estratégicas.

## Funcionalidades

- Leitura de dados de um arquivo Excel contendo informações de vendas.
- Limpeza e preparação dos dados para análise.
- Geração de gráficos utilizando a biblioteca Matplotlib para visualização dos dados.
- Agrupamento e cálculo de métricas como média e soma para diferentes categorias (produto, região, equipe de venda, método de pagamento).
- Envio automatizado de e-mail com os gráficos gerados como anexos.

## Stacks Utilizadas

- Python
- Pandas
- Matplotlib
- openpyxl
- smtplib (para envio de e-mails)
- email.mime (para construção do e-mail)

## Instalação

Para utilizar este projeto localmente, siga os passos abaixo:

### Como clonar o repositório

```bash
git clone https://github.com/seu_usuario/nome-do-repositorio.git
```

### Como ativar o ambiente virtual 
```bash
cd jao-projeto-analise-de-vendas
python -m venv venv
```
### No Windows
```bash
venv\Scripts\activate
```
### No Linux/Mac
```bash
source venv/bin/activate
```
