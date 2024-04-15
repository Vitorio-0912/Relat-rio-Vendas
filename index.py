import pandas as pd
import win32com.client as win32
# importando a base de dados
tabela_vendas = pd.read_excel('Vendas.xlsx')


# visualizar a base de dados
pd.set_option('display.max_columns', None)


# faturamento por loja
faturamento = tabela_vendas [['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)

# quantidade de produtos vendidos por loja
qtd_produtos = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(qtd_produtos) 

print('-' * 50)

# ticket medio por produto em cada loja
ticket_medio = (faturamento['Valor Final'] / qtd_produtos['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Medio'})
print(ticket_medio)


# enviar um email com o relatorio  
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'vitoriog616@gmail.com'
mail.Subject = 'Relatório de Vendas por Loja'
mail.HTMLBody = f'''
<p>Prezados,</p>

<p>Segue o Relatório de Vendas por cada Loja.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:.2f}'.format})} 

<p>Quantidade Vendida:</p>
{qtd_produtos.to_html()}

<p>Ticket Médio dos Produtos em cada Loja:</p>
{ticket_medio.to_html(formatters={'Ticket Medio': 'R${:.2f}'.format})}

<p>Qualquer duvida estou a disposição.</p>

<p>att..</p>
<p>Vitorio</p>

'''

mail.Send()

print('Email Enviado')
