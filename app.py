#libs
import os
import pathlib
import pandas as pd
from time import sleep
import win32com.client as win32


# - importar e tratar as bases de dados
# juntar o vendas com loja e criar 1 arquivo para cada loja
# Salvar o backup nas pastas
# Calcular os indicadores
# Enviar o OnePage
# Enviar email p/ diretoria
# Salvar o Backup

# ETL - extração, transformação, carregamento



emails = pd.read_excel(r'C:\PYTHON\Projeto AutomacaoIndicadores\Bases de Dados\Emails.xlsx')
lojas = pd.read_csv(r'C:\PYTHON\Projeto AutomacaoIndicadores\Bases de Dados\Lojas.csv', encoding='latin1', sep=";")
vendas = pd.read_excel(r'C:\PYTHON\Projeto AutomacaoIndicadores\Bases de Dados\Vendas.xlsx')
#print(emails)
#print(lojas)
#print(vendas)

# incluir nome da loja em vendas

vendas = vendas.merge(lojas, on='ID Loja') #mesclar as duas tabelas, mas, as duas colunas ID e LOJA
# print(vendas)


#formular o dataframe para cada loja e salvar em um dicionário
dicionario_lojas = {}
for loja in lojas['Loja']: # para cada uma das lojas na coluna que há o nome da loja no dataframe

    dicionario_lojas[loja] = vendas.loc[vendas['Loja']==loja, :]
# print(dicionario_lojas['Rio Mar Recife'])
# print(dicionario_lojas['Shopping Vila Velha'])


#pegar a data 
dia_indicador = vendas['Data'].max() # formato Datetime
# print(f'{dia_indicador.day}/{dia_indicador.month}/{dia_indicador.year}') #formatando a data

#identificar se a pasta já existe
caminho_backup = pathlib.Path("Backup Arquivos Lojas")

arquivos_pasta_backup = caminho_backup.iterdir()
lista_nomes_backup = [arquivo.name for arquivo in arquivos_pasta_backup] # list compreension

#criar pastas das lojas
# for loja in dicionario_lojas:
    
#     if loja not in lista_nomes_backup:
#         nova_pasta = caminho_backup / loja
#         nova_pasta.mkdir()
#         sleep(1)

#     #salvar dentro da pasta
#     nome_arquivo = '{}_{}_{}.xlsx'.format(dia_indicador.month, dia_indicador.day, loja)
#     local_arquivo = caminho_backup / loja / nome_arquivo
#     dicionario_lojas[loja].to_excel(local_arquivo)
#     sleep(1)



loja = 'Norte Shopping'
vendas_loja = (dicionario_lojas[loja])
vendas_loja_dia = vendas_loja.loc[vendas_loja['Data']==dia_indicador, :]


#faturamento
#faturamento do ano somado
faturamento_ano = vendas_loja['Valor Final'].sum()
#print(faturamento_ano)
#faturamento do dia indicado
faturamento_dia = vendas_loja_dia['Valor Final'].sum()
#print(faturamento_dia)
#faturamento

#diversidade de produtos
qtde_produtos_ano = len(vendas_loja['Produto'].unique()) # o unique pega os valores da coluna, mas ele tira os duplicados !
qtde_produtos_dia = len(vendas_loja_dia['Produto'].unique()) # o unique pega os valores da coluna, mas ele tira os duplicados !

#print(qtde_produtos_ano)
#print(qtde_produtos_dia)

#ticket médio - 
valor_venda = vendas_loja.groupby('Código Venda').sum() # groupby vai agrupar utilizando-se da coluna informada, no caso "código venda"
ticket_medio_ano = valor_venda['Valor Final'].mean() # o mean vai retonar a média
#print(ticket_medio_ano)

#ticket_medio_dia = 
valor_venda_dia = vendas_loja_dia.groupby('Código Venda').sum() # groupby vai agrupar utilizando-se da coluna informada, no caso "código venda"
ticket_medio_dia = valor_venda_dia['Valor Final'].mean() # o mean vai retonar a média
#print(ticket_medio_dia)

#definição de metas

meta_faturamento_dia = 1000
meta_faturamento_ano = 1650000
meta_qtdeprodutos_dia = 4
meta_qtdeprodutos_ano = 120
meta_ticketmedio_dia = 500
meta_ticketmedio_ano = 500


# # Criação do envio de e-mail
outlook = win32.Dispatch('outlook.application')

nome = emails.loc[emails['Loja']==loja,'Gerente'].values[0]
mail = outlook.CreateItem(0)
mail.To = 'pedro.leon23@outlook.com.br'#emails.loc[emails['Loja']==loja,'E-mail'].values[0] #loc['linha','coluna'] o values[0] vai retornar somente o valor do email
mail.CC = '' # cópia
mail.BCC = '' #cópia oculta
mail.Subject = f'OnePage dia {dia_indicador.day}/{dia_indicador.month} - Loja {loja}' #assunto
mail.Body = ''
#ou mail.HTMKBody = '<p>Corpo do Email em HTML </p>'

# Anexos (pode-se colocar quantos quiser):
attachmet = pathlib.Path.cwd() / caminho_backup / loja / f'{dia_indicador.month}_{dia_indicador.day}_{loja}.xlsx'
mail.attachments.Add(str(attachmet))

mail.Send()
