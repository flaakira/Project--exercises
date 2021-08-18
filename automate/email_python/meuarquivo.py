#importar a base de dados
import pandas as pd
from pandas.core.groupby.base import OutputKey
import win32com.client as win32

vendas = pd.read_excel('Vendas.xlsx')

#visualizar a base de dados

pd.set_option('display.max_columns', None)
print(vendas)
#Faturamento por loja

faturamento = vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)
#Quantidade de produto vendido por loja

quantidade = vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade)
#ticket medio por produto.

media = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
media = media.rename(columns={0: 'Ticket Medio'})
print(media)
#enviar um email com o relatorio

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'FOliveira@cervusequipment.com'
mail.Subject = 'Relatorio de vendas por Loja'
mail.HTMLBody = f'''
<p>Prezados,</p>

<p>Segue o Relatorio de vendas por cada Loja.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade Vendida:</p>
{quantidade.to_html()}

<p>Ticket Medio dos produtos em cada Loja:</p>
{media.to_html(formatters={'Ticket Medio': 'R${:,.2f}'.format})}

<p>Qualquer duvida estou a disposicao</p>

<p>Att,</p>
<p>Flavio Akira</p>
'''

mail.Send()

