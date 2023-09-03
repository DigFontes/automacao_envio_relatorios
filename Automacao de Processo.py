#!/usr/bin/env python
# coding: utf-8

# ### Passo 1 - Importar Arquivos e Bibliotecas

# In[11]:


#importar bibliotecas
import pandas as pd
import win32com.client as win32
import openpyxl 


# In[12]:


#importando base de dados, lendo arquivos no formato excel e csv
emails_df = pd.read_excel(r'C:\Users\Virtual Office\Python\Módulo 37 - Projeto 1\Bases de Dados\Emails.xlsx')
lojas_df = pd.read_csv(r'C:\Users\Virtual Office\Python\Módulo 37 - Projeto 1\Bases de Dados\Lojas.csv', encoding = 'latin1',sep = ';')
vendas_df = pd.read_excel(r'C:\Users\Virtual Office\Python\Módulo 37 - Projeto 1\Bases de Dados\Vendas.xlsx')


# ### Passo 2 - Definir Criar uma Tabela para cada Loja e Definir o dia do Indicador

# In[13]:


vendas_df = vendas_df.merge(lojas_df, on = 'ID Loja')
print('Base de dados importada para o programa!')


# In[14]:


dicionario_lojas = {}
for loja in lojas_df['Loja']:
    dicionario_lojas[loja] = vendas_df.loc[vendas_df['Loja']==loja,:]


# In[15]:


dia_indicador = vendas_df['Data'].max()


# In[16]:


print('Tratamentos na base de dados concluída! Pronto para salvar os arquivos de cada loja.')


# ### Passo 3 - Salvar a planilha na pasta de backup

# In[17]:


#identificar se a pasta já existe
import pathlib
caminho_backup = pathlib.Path(r'C:\Users\Virtual Office\Python\Módulo 37 - Projeto 1\Backup Arquivos Lojas')
arquivos_pasta_backup = caminho_backup.iterdir()

lista_nomes_backup = [arquivo.name for arquivo in arquivos_pasta_backup]

for loja in dicionario_lojas:
    if loja not in lista_nomes_backup:
        nova_pasta = caminho_backup / loja
        nova_pasta.mkdir()
    #salvar dentro da pasta   
    nome_arquivo = '{}_{}_{}.xlsx'.format(dia_indicador.month, dia_indicador.day,loja)    
    local_arquivo = caminho_backup / loja / nome_arquivo
    
    dicionario_lojas[loja].to_excel(local_arquivo)

    print(nome_arquivo)
    print('Arquivos salvos!')


# ### Passo 4 - Calcular o indicador para 1 loja

# In[18]:


#definições da meta
meta_faturamento_dia = 1000
meta_faturamento_ano = 1650000
meta_qtdeprodutos_dia = 4
meta_qtdeprodutos_ano = 120
meta_ticketmedio_dia = 500
meta_ticketmedio_ano = 2200


# In[19]:


print('Pronto para enviar os e-mails para os gerentes!')


# In[20]:


for loja in dicionario_lojas:
    
    vendas_loja = dicionario_lojas[loja]
    vendas_loja_dia = vendas_loja.loc[vendas_loja['Data'] == dia_indicador,:]
    #faturamento
    faturamento_ano = vendas_loja['Valor Final'].sum()
    faturamento_dia = vendas_loja_dia['Valor Final'].sum()
    #Diversidade de produtos
    qtde_produtos_ano = len(vendas_loja['Produto'].unique())
    qtde_produtos_dia = len(vendas_loja_dia['Produto'].unique())
    #ticket medio
    valor_venda = vendas_loja[['Código Venda','ID Loja', 'Quantidade', 'Valor Unitário', 'Valor Final']]
    valor_venda = valor_venda.groupby('Código Venda').sum()
    ticket_medio_ano = valor_venda['Valor Final'].mean()

    #ticket medio dia
    valor_venda_dia = vendas_loja_dia[['Código Venda','ID Loja', 'Quantidade', 'Valor Unitário', 'Valor Final']]
    valor_venda_dia = valor_venda_dia.groupby('Código Venda').sum()
    ticket_medio_dia = valor_venda_dia['Valor Final'].mean()

    
    print(faturamento_ano, faturamento_dia, qtde_produtos_ano, qtde_produtos_dia, ticket_medio_ano, ticket_medio_dia)
    
    
   
    outlook = win32.Dispatch('outlook.application')

    nome = emails_df.loc[emails_df['Loja'] == loja, 'Gerente'].values[0]

    mail = outlook.CreateItem(0)
    mail.To = emails_df.loc[emails_df['Loja'] == loja, 'E-mail'].values[0] 
    mail.Subject = 'OnePage Dia {}/{} - Loja {}'.format(dia_indicador.day, dia_indicador.month, loja)
    #Corpo do email

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
    <p> Bom dia, {nome}</p>

    <p>O resultado de ontem <strong> ({dia_indicador.day}/{dia_indicador.month})</strong> da loja <strong>{loja}</strong>foi: </p>

    <table>
      <tr>
        <th>Indicador</th>
        <th>Valor Dia</th>
        <th>Meta Dia</th>
        <th>Cenário Dia</th>
      </tr>
      <tr>
        <td>Faturamento</td>
        <td>R${faturamento_dia:.2f}</td>
        <td>R${meta_faturamento_dia:.2f}</td>
        <td><font color = '{cor_fat_dia}'>◙</font></td>
      </tr>
      <tr>
        <td>Diversidade de Produtos</td>
        <td>{qtde_produtos_dia}</td>
        <td>{meta_qtdeprodutos_dia}</td>
        <td><font color = '{cor_qtde_dia}'>◙</font</td>
      </tr>
      <tr>
         <td>Ticket Médio</td>
        <td>R${ticket_medio_dia:.2f}</td>
        <td>R${meta_ticketmedio_dia:.2f}</td>
        <td><font color = '{cor_ticket_dia}'>◙</font</td>
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
        <td>R${faturamento_ano:.2f}</td>
        <td>R${meta_faturamento_ano:.2f}</td>
        <td><font color = '{cor_fat_ano}'>◙</font></td>
      </tr>
      <tr>
        <td>Diversidade de Produtos</td>
        <td>{qtde_produtos_ano}</td>
        <td>{meta_qtdeprodutos_ano}</td>
        <td><font color = '{cor_qtde_ano}'>◙</font</td>
      </tr>
      <tr>
         <td>Ticket Médio</td>
        <td>R${ticket_medio_ano:.2f}</td>
        <td>R${meta_ticketmedio_ano:.2f}</td>
        <td><font color = '{cor_ticket_ano}'>◙</font</td>
      </tr>
    </table>
    <p>Segue em anexo a planilha com todos os dados para mais detalhes.</p>

    <p>Qualquer dúvida estou à disposição.</p>
    <p> Att., Diêgo Fontes</p>
    <br>

    </br>
    '''
    #Anexos 
    attachment = pathlib.Path.cwd() / caminho_backup / loja / f'{dia_indicador.month}_{dia_indicador.day}_{loja}.xlsx'
    mail.Attachments.Add(str(attachment))

    mail.Send()
    print('E-mail da Loja {} enviado.'.format(loja))


# In[21]:


print('Envio dos emails concluída! Confira na caixa de saída.')


# ### Passo 5 - Enviar por e-mail para o gerente

# In[ ]:





# ### Passo 6 - Automatizar todas as lojas

# ### Passo 7 - Criar ranking para diretoria

# In[22]:


faturamento_lojas = vendas_df[['Loja', 'Valor Final']]

faturamento_lojas = faturamento_lojas.groupby('Loja').sum()
faturamento_lojas_ano = faturamento_lojas.sort_values(by = 'Valor Final', ascending= False)

nome_arquivo = '{}_{}_Ranking Anual.xlsx'.format(dia_indicador.month, dia_indicador.day)
faturamento_lojas_ano.to_excel(r'C:\Users\Virtual Office\Python\Módulo 37 - Projeto 1\Backup Arquivos Lojas\{}'.format(nome_arquivo))

vendas_dia = vendas_df.loc[vendas_df['Data']==dia_indicador, :]
faturamento_lojas_dia = vendas_dia[['Loja', 'Valor Final']]                         
faturamento_lojas_dia = faturamento_lojas_dia.groupby('Loja').sum()
faturamento_lojas_dia = faturamento_lojas_dia.sort_values(by = 'Valor Final', ascending = False)

nome_arquivo = '{}_{}_Ranking Dia.xlsx'.format(dia_indicador.month, dia_indicador.day)
faturamento_lojas_dia.to_excel(r'C:\Users\Virtual Office\Python\Módulo 37 - Projeto 1\Backup Arquivos Lojas\{}'.format(nome_arquivo))


# ### Passo 8 - Enviar e-mail para diretoria

# In[23]:


outlook = win32.Dispatch('outlook.application')

nome = emails_df.loc[emails_df['Loja'] == loja, 'Gerente'].values[0]

mail = outlook.CreateItem(0)
mail.To = emails_df.loc[emails_df['Loja'] == 'Diretoria', 'E-mail'].values[0] 
mail.Subject = 'Ranking Dia'.format(dia_indicador.day, dia_indicador.month, loja)
#Corpo do email
mail.Body = f'''
Prezados, 

Melhor loja do dia em faturamento: Loja {faturamento_lojas_dia.index[0]} com faturamento R${faturamento_lojas_dia.iloc[0,0]:.2f}
Pior loja do dia em faturamento: Loja {faturamento_lojas_dia.index[-1]} com faturamento R${faturamento_lojas_dia.iloc[-1,0]:.2f}

Melhor loja do ano em faturamento: Loja {faturamento_lojas_ano.index[0]} com faturamento R${faturamento_lojas_ano.iloc[0,0]:.2f}
Pior loja do ano em faturamento: Loja {faturamento_lojas_ano.index[-1]} com faturamento R${faturamento_lojas_ano.iloc[-1,0]:.2f}

Segue em anexo raqueamento diário e anual de todas as lojas.

Qualquer dúvida estou a disposição.

Att,
Diêgo Fontes

'''
#Anexos 
attachment = pathlib.Path.cwd() / caminho_backup /f'{dia_indicador.month}_{dia_indicador.day}_Ranking Anual.xlsx'
mail.Attachments.Add(str(attachment))
attachment = pathlib.Path.cwd() / caminho_backup /f'{dia_indicador.month}_{dia_indicador.day}_Ranking Dia.xlsx'
mail.Attachments.Add(str(attachment))

mail.Send()
print('E-mail da Diretoria enviado')
print('Programa concluído!')

