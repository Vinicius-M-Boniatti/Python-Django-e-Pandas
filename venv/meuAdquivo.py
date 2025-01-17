import pandas as pd



#importar a base de dados

tabela_de_vendas = pd.read_excel('vendas.xlsx')
pd.set_option('display.max_columns', None)
print(tabela_de_vendas)


#visualizar a base de dados



#faturamento da loja
faturamento = tabela_de_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)
print("="*20)

#quantidade de proputos
print("QUANTIDADE")
print("="*20)
quantidade = tabela_de_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade)
print("="*20)

#tiket medio por loja

print("TICKET MÉDIO")
print("="*20)
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
print(ticket_medio)

#enviar um email
import win32com.client as win32

outlook = win32.Dispatch('outlook.application')
mail = outlook.createItem(0)
mail.To = 'viniciusmboniatti@gmail.com'
mail.Subject = 'Manssage Subject'
mail.HTMLBody ='<h2> MENSAGEM QUE SERÁ ENVIADA </h2>'
mail.send()

# ele envia a mensagem pro email descrito em cima, usando o 
# outlook enviando pra qual email, tem que estar logado no app 
# do outlook no PC.
 



