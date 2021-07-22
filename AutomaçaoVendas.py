import pandas as pd  # Biblioteca q "cria" um excel no python , usando tabelas
import win32com.client as win32
import time

print()
print('\033[1;30m-=\033[m' * 25)
print('\033[1;31m AUTOMAÇAO - ENVIO DO RELATORIO DE VENDAS POR LOJA \033[m'.center(50))
print('\033[1;30m-=\033[m' * 25)

print('\033[36m Processando... \033[m')
time.sleep(1)
print('\033[1;30m-=\033[m' * 45)

tabela_vendas = pd.read_excel('Vendas.xlsx')  # Importa a base de dados do excel
pd.set_option('display.max_columns', None)  # Mostrar o maximo de colunas do arquivo - Visualizando a base de dados
print(tabela_vendas)

# Faturamento por loja - Filtra primeiro. Agrupando todas as lojas e somando o faturamento total de cada uma delas
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)
print('\033[1;30m-=\033[m' * 45)

# Quantidade de produtos vendidos por lojas
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade)
print('\033[1;30m-=\033[m' * 45)

# Ticket médio por produto em cada loja ( faturamento divido por quantidade de produtos )
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()  # To_frame = Transformando em tabela
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})  # Mudando o nome da coluna da tabela de ticket médio
print(ticket_medio)

# Enviar um email com o relatorio
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.to = 'xxxxxxxxxxxx'  # O email do destinatario
mail.Subject = 'Relatório de Vendas por Loja'  # Titulo do email
mail.HTMLBody = f'''
<p>Prezados,</p>

<p>Segue o Relatório de Vendas por cada loja.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})} 

<p>Quantidade Vendida:</p>
{quantidade.to_html()}

<p>Ticket Médio dos produtos em cada loja:</p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Qualquer dúvida estou a disposição.</p>

<p>Att.,</p>
<p>xxxxxxxxxxxxxxx</p>
'''

mail.Send()

print('\033[1;30m-=\033[m' * 45)
print('\033[1;31m E-mail enviado com sucesso! \033[m')
