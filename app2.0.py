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
# criar ranking das lojas pq diretoria
# Salvar o Backup

# ETL - extração, transformação, carregamento

# importar as bases de dados

emails = pd.read_excel(r'Bases de Dados\Emails.xlsx')
lojas = pd.read_csv(r'Bases de Dados\Lojas.csv', encoding='latin1', sep=';')
vendas = pd.read_excel(r'Bases de Dados\Vendas.xlsx')

############################

# juntar a lista de lojas com a lista de vendas

vendas = vendas.merge(lojas, on='ID Loja')


# Cabeçalho do arquivo final:
# Código Venda | Data | ID Loja | Produto | Quantidade | Valor Unitário | Valor Final | Loja

###############################################

# formular o dataframe para cada loja e salvar todas as lojas em um dicionaro => dicionario_lojas

dicionario_lojas = {}
for loja in lojas['Loja']:
    dicionario_lojas[loja] = vendas.loc[vendas['Loja'] ==loja, :]

# com o comando => print(dicionario_lojas['Rio Mar Recife']) , 
# é possível retornar uma tabela somente da loja mencionada, já que todas estão no dicionario_lojas

##################################################################################################

# pegar a data da tabela

dia_indicador = vendas['Data'].max() # print(f'{dia_indicador.day}/{dia_indicador.month}') #formatado a data

########################

# salvar as planilhas na pasta de backup

# identificar se a pasta da loja já existe

caminho_backup = pathlib.Path(r'Backup Arquivos Lojas')

arquivos_pasta_backup = caminho_backup.iterdir()
lista_nome_backup = [arquivo.name for arquivo in arquivos_pasta_backup] 


# com o iterdir() e um for, é possível listar todos os arquivos/diretórios dentro do caminho. print(lista_nome_backup)
lista_pastas = []
for loja in dicionario_lojas:
    if loja not in lista_nome_backup:
        nova_pasta = caminho_backup / loja
        nova_pasta.mkdir() # cria-se a pasta no caminho do backup + nome da loja
        lista_pastas.append(loja)

    

# salvar a planilha da loja dentro de cada pasta

        nome_arquivo = f'{dia_indicador.month}_{dia_indicador.day}_{loja}.xlsx' # {}_{}_{}.xlsx
        local_arquivo = caminho_backup / loja / nome_arquivo

        dicionario_lojas[loja].to_excel(local_arquivo) 

# este comando é... sem comentários... recolher a planilha do dicionário de lojas e transforma-lo em excel

if lista_pastas:
    print(f'foram criados as pastas {lista_pastas}.')
else:
    print('Não foram criadas pastas. Prosseguindo...')

# definição de metas

meta_faturamento_dia = 1000
meta_faturamento_ano = 1650000
meta_qtdeprodutos_dia = 4
meta_qtdeprodutos_ano = 120
meta_ticketmedio_dia = 500
meta_ticketmedio_ano = 500

##########################
#loja = 'Norte Shopping'
for loja in dicionario_lojas:

    vendas_loja = dicionario_lojas[loja]
    vendas_loja_dia = vendas_loja.loc[vendas_loja['Data']==dia_indicador, :]

    # faturamento
    faturamento_ano = vendas_loja['Valor Final'].sum()
    faturamento_dia = vendas_loja_dia['Valor Final'].sum()

    # diversidade de produtos
    qtde_produtos_ano = len(vendas_loja['Produto'].unique()) # unique é para valores únicos, len para quantidade numérica
    qtde_produtos_dia = len(vendas_loja_dia['Produto'].unique())

    #ticket médio
    valor_venda = vendas_loja.groupby('Código Venda').sum()
    ticket_medio_ano = valor_venda['Valor Final'].mean() # mean => média

    valor_venda_dia = vendas_loja_dia.groupby('Código Venda').sum()
    ticket_medio_dia = valor_venda_dia['Valor Final'].mean()

    #teste de print para validar se as informações estão retornando corretamente

    # print(f'faturamento anual {loja}: R${faturamento_ano}')
    # print(f'faturamento diário {dia_indicador.day}/{dia_indicador.month} {loja}: R${faturamento_dia}')
    # print('###')
    # print(f'quantidade de produtos vendidos anual  {loja}: {qtde_produtos_ano}')
    # print(f'quantidade de produtos vendidos diário {dia_indicador.day}/{dia_indicador.month} {loja}: {qtde_produtos_dia}')
    # print('###')
    # print(f'ticket médio anual  {loja}: R${ticket_medio_ano}')
    # print(f'ticket médio diário {dia_indicador.day}/{dia_indicador.month} {loja}: R${ticket_medio_dia}')

    ############################################################################

    # criação do envio do e-mail

    outlook = win32.Dispatch('outlook.application')

    nome = emails.loc[emails['Loja']==loja, 'Gerente'].values[0]
    mail = outlook.CreateItem(0)
    mail.To = 'pedroleonpython@gmail.com' #emails.loc[emails['Loja']==loja, 'E-mail'].values[0]
    mail.Subject = f'OnePage Dia {dia_indicador.day}/{dia_indicador.month} - Loja {loja}'
    #mail.Body = 'Texto do E-mail'

    if faturamento_dia >= meta_faturamento_dia:
            cor_fat_dia = 'green'
    else:
        cor_fat_dia = 'red'
    if faturamento_ano >= meta_faturamento_ano:
        cor_fat_ano = 'green'
    else:
        cor_fat_ano = 'red'
    if qtde_produtos_dia >= meta_qtdeprodutos_dia:
        cor_qtde_dia = 'green'
    else:
        cor_qtde_dia = 'red'
    if qtde_produtos_ano >= meta_qtdeprodutos_ano:
        cor_qtde_ano = 'green'
    else:
        cor_qtde_ano = 'red'
    if ticket_medio_dia >= meta_ticketmedio_dia:
        cor_ticket_dia = 'green'
    else:
        cor_ticket_dia = 'red'
    if ticket_medio_ano >= meta_ticketmedio_ano:
        cor_ticket_ano = 'green'
    else:
        cor_ticket_ano = 'red'

    mail.HTMLBody = f'''
    <p>Bom dia, {nome}</p>

    <p>O resultado de ontem <strong>({dia_indicador.day}/{dia_indicador.month})</strong> da <strong>Loja {loja}</strong> foi:</p>

    <table>
        <tr>
        <th>Indicador</th>
        <th>Valor Dia</th>
        <th>Meta Dia</th>
        <th>Cenário Dia</th>
        </tr>
        <tr>
        <td>Faturamento</td>
        <td style="text-align: center">R${faturamento_dia:.2f}</td>
        <td style="text-align: center">R${meta_faturamento_dia:.2f}</td>
        <td style="text-align: center"><font color="{cor_fat_dia}">◙</font></td>
        </tr>
        <tr>
        <td>Diversidade de Produtos</td>
        <td style="text-align: center">{qtde_produtos_dia}</td>
        <td style="text-align: center">{meta_qtdeprodutos_dia}</td>
        <td style="text-align: center"><font color="{cor_qtde_dia}">◙</font></td>
        </tr>
        <tr>
        <td>Ticket Médio</td>
        <td style="text-align: center">R${ticket_medio_dia:.2f}</td>
        <td style="text-align: center">R${meta_ticketmedio_dia:.2f}</td>
        <td style="text-align: center"><font color="{cor_ticket_dia}">◙</font></td>
        </tr>
    </table>
    <br>
    <table>
        <tr>
        <th>Indicador</th>
        <th>Valor Ano</th>
        <th>Meta Ano</th>
        <th>Cenário Ano</th>
        </tr>
        <tr>
        <td>Faturamento</td>
        <td style="text-align: center">R${faturamento_ano:.2f}</td>
        <td style="text-align: center">R${meta_faturamento_ano:.2f}</td>
        <td style="text-align: center"><font color="{cor_fat_ano}">◙</font></td>
        </tr>
        <tr>
        <td>Diversidade de Produtos</td>
        <td style="text-align: center">{qtde_produtos_ano}</td>
        <td style="text-align: center">{meta_qtdeprodutos_ano}</td>
        <td style="text-align: center"><font color="{cor_qtde_ano}">◙</font></td>
        </tr>
        <tr>
        <td>Ticket Médio</td>
        <td style="text-align: center">R${ticket_medio_ano:.2f}</td>
        <td style="text-align: center">R${meta_ticketmedio_ano:.2f}</td>
        <td style="text-align: center"><font color="{cor_ticket_ano}">◙</font></td>
        </tr>
    </table>

    <p>Segue em anexo a planilha com todos os dados para mais detalhes.</p>

    <p>Qualquer dúvida estou à disposição.</p>
    <p>Att., Leon - Análista Python</p>
    '''

    # Anexos (pode colocar quantos quiser):
    attachment  = pathlib.Path.cwd() / caminho_backup / loja / f'{dia_indicador.month}_{dia_indicador.day}_{loja}.xlsx'
    mail.Attachments.Add(str(attachment))

    mail.Send()
    sleep(30)
    print(f'E-mail da Loja {loja} enviado')
