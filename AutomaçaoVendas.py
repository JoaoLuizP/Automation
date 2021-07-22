import pandas as pd  # Library that "creates" an excel in python , using tables
import win32com.client as win32
import time

print()
print('\033[1;30m-=\033[m' * 25)
print('\033[1;31m AUTOMAÇAO - ENVIO DO RELATORIO DE VENDAS POR LOJA \033[m'.center(50))
print('\033[1;30m-=\033[m' * 25)

print('\033[36m Processando... \033[m')
time.sleep(1)
print('\033[1;30m-=\033[m' * 45)

tabela_vendas = pd.read_excel('Vendas.xlsx')  # Import to excel database
pd.set_option('display.max_columns', None)  # Show maximum columns of file - Viewing database
print(tabela_vendas)

# Billing by Store - Filters first. Grouping all stores and adding up the total sales of each one of them
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)
print('\033[1;30m-=\033[m' * 45)

# Number of products sold by stores
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade)
print('\033[1;30m-=\033[m' * 45)

# Average ticket per product in each store (invoicing divided by quantity of products)
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()  # To_frame = Turning into table
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})  # Changing the Column Name of the Average Ticket Table
print(ticket_medio)

# Send an email with the report
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.to = 'xxxxxxx@hotmail.com'  # The recipient's email
mail.Subject = 'Relatório de Vendas por Loja'  # Email title
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
<p> xxxxxxx </p>
'''

mail.Send()

print('\033[1;30m-=\033[m' * 45)
print('\033[1;31m E-mail enviado com sucesso! \033[m')
