import pandas as pd



#importar a base de dados

tabela_de_vendas = pd.read_excel('vendas.xlsx')
pd.set_option('display.max_columns', None)
print(tabela_de_vendas)


#visualizar a base de dados



#faturamento da loja
faturamento = tabela_de_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)
print("=========================================")

#quantidade de proputos
print("QUANTIDADE")
quantidade = tabela_de_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade)
print("=========================================")

#tiket medio por loja

print("TICKET MÉDIO")
print("=========================================")
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
print(ticket_medio)

#enviar um email



#visualizar a base de dados