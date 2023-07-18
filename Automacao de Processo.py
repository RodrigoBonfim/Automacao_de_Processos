#!/usr/bin/env python
# coding: utf-8

# ### Passo 1 - Importar Arquivos e Bibliotecas

# In[1]:


import pandas as pd
import pathlib 
import win32com.client as win32


# In[2]:


emails = pd.read_excel(r'Bases de dados\Emails.xlsx')
lojas = pd.read_csv(r'Bases de dados\Lojas.csv', encoding='latin1', sep=';') 
vendas = pd.read_excel(r'Bases de dados\Vendas.xlsx')
display(emails)
display(lojas)
display(vendas)


# ### Passo 2 - Definir Criar uma Tabela para cada Loja e Definir o dia do Indicador

# In[3]:


#incluir o nome da loja em vendas
vendas = vendas.merge(lojas, on='ID Loja')
display(vendas)


# In[4]:


dicionario_lojas = {}
for loja in lojas ['Loja']:
    dicionario_lojas[loja] = vendas.loc[vendas['Loja']==loja, :]
display(dicionario_lojas['Salvador Shopping'])
display(dicionario_lojas['Shopping Eldorado'])


# In[5]:


dia_indicador = vendas['Data'].max()
print(dia_indicador)


# ### Passo 3 - Salvar a planilha na pasta de backup

# In[6]:


#identificar se a pasta já existe
caminho_backup = pathlib.Path(r'Backup Arquivos Lojas')

arquivos_backup = caminho_backup.iterdir()
lista_arquivos = [arquivo.name for arquivo in arquivos_backup]

for loja in dicionario_lojas:
    if loja not in lista_arquivos:
        nova_pasta = caminho_backup/loja
        nova_pasta.mkdir()
    #salvar dentro da pasta
    nome_arquivo = '{}_{}_{}.xlsx'.format(dia_indicador.month, dia_indicador.day, loja)
    local_arquivo = caminho_backup/loja/nome_arquivo
    dicionario_lojas[loja].to_excel(local_arquivo)
    


# ### Passo 4 - Calcular o indicador para 1 loja e automatizar para as demais
# - Enviar E-mail para todos os gerentes sobre suas lojas.

# In[7]:


#Definição de metas

meta_faturamento_dia = 1000
meta_faturamento_ano = 1650000
meta_qtde_produtos_dia = 4
meta_qtde_produtos_ano = 120
meta_ticket_medio_dia = 500
meta_ticket_medio_ano = 500


# In[8]:


for loja in dicionario_lojas:
    vendas_loja = dicionario_lojas[loja]
    vendas_loja_dia = vendas_loja.loc[vendas_loja['Data']==dia_indicador, :]

    #####faturamento
    faturamento_ano = vendas_loja['Valor Final'].sum()
    #print(faturamento_ano)
    faturamento_dia = vendas_loja_dia['Valor Final'].sum()
    #print(faturamento_dia)
    ######diversidade de produtos
    quantidade_produtos_ano = len(vendas_loja['Produto'].unique())
    #print(quantidade_produtos_ano)
    quantidade_produtos_dia= len(vendas_loja_dia['Produto'].unique())
    #print(quantidade_produtos_dia)
    ######ticket Médio
    valor_venda = vendas_loja.groupby('Código Venda').sum()
    ticket_medio_ano = valor_venda['Valor Final'].mean()
    #print(ticket_medio_ano)
    valor_venda_dia = vendas_loja_dia.groupby('Código Venda').sum()
    ticket_medio_dia = valor_venda_dia['Valor Final'].mean()
    #print(ticket_medio_dia)
    outlook = win32.Dispatch('outlook.application')

    nome = emails.loc[emails['Loja']==loja, 'Gerente'].values[0]
    mail = outlook.CreateItem(0)
    mail.To = emails.loc[emails['Loja']==loja, 'E-mail'].values[0]
    mail.Subject = f'OnePage Dia {dia_indicador.day}/{dia_indicador.month} - Loja {loja}'
   
    if faturamento_dia >= meta_faturamento_dia:
        cor_fat_dia = 'green'
    else:
        cor_fat_dia = 'red'
    if faturamento_ano >= meta_faturamento_ano:
        cor_fat_ano = 'green'
    else:
        cor_fat_ano = 'red'
    if quantidade_produtos_dia >= meta_qtde_produtos_dia:
        cor_qtde_dia = 'green'
    else:
        cor_qtde_dia = 'red'
    if quantidade_produtos_ano >= meta_qtde_produtos_ano:
        cor_qtde_ano = 'green'
    else:
        cor_qtde_ano = 'red'
    if ticket_medio_dia >= meta_ticket_medio_dia:
        cor_ticket_dia = 'green'
    else:
        cor_ticket_dia = 'red'
    if ticket_medio_ano >= meta_ticket_medio_ano:
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
            <td style="text-align: center">{quantidade_produtos_dia}</td>
            <td style="text-align: center">{meta_qtde_produtos_dia}</td>
            <td style="text-align: center"><font color="{cor_qtde_dia}">◙</font></td>
          </tr>
          <tr>
            <td>Ticket Médio</td>
            <td style="text-align: center">R${ticket_medio_dia:.2f}</td>
            <td style="text-align: center">R${meta_ticket_medio_dia:.2f}</td>
            <td style="text-align: center"><font color="{cor_ticket_dia}">◙</font></td>
          </tr>
        </table>
        <br>
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
            <td style="text-align: center">{quantidade_produtos_ano}</td>
            <td style="text-align: center">{meta_qtde_produtos_ano}</td>
            <td style="text-align: center"><font color="{cor_qtde_ano}">◙</font></td>
          </tr>
          <tr>
            <td>Ticket Médio</td>
            <td style="text-align: center">R${ticket_medio_ano:.2f}</td>
            <td style="text-align: center">R${meta_ticket_medio_ano:.2f}</td>
            <td style="text-align: center"><font color="{cor_ticket_ano}">◙</font></td>
          </tr>
        </table>

        <p>Segue em anexo a planilha com todos os dados para mais detalhes.</p>

        <p>Qualquer dúvida estou à disposição.</p>
        <p>Atenciosamente, Rodrigo Bonfim.</p>
        '''
    #Anexos
    attachment = pathlib.Path.cwd()/caminho_backup/loja/f'{dia_indicador.month}_{dia_indicador.day}_{loja}.xlsx'

    mail.Attachments.Add(str(attachment))

    mail.Send()
    print('E-mail da Loja {} enviado'.format(loja))


# ### Passo 5 - Criar ranking para diretoria

# In[9]:


faturamento_loja = vendas.groupby('Loja')[['Loja', 'Valor Final']].sum()
faturamento_loja_ano = faturamento_loja.sort_values(by='Valor Final', ascending=False)
display(faturamento_loja)

nome_arquivo = '{}_{}_Ranking Anual.xlsx'.format(dia_indicador.month, dia_indicador.day)
faturamento_loja_ano.to_excel(r'Backup Arquivos Lojas\{}'.format(nome_arquivo))

vendas_dia = vendas.loc[vendas["Data"]==dia_indicador,:]
faturamento_loja_dia = vendas_dia.groupby('Loja')[['Loja', 'Valor Final']].sum()
faturamento_loja_dia = faturamento_loja_dia.sort_values(by='Valor Final', ascending=False)
display(faturamento_loja_dia)

nome_arquivo = '{}_{}_Ranking Dia.xlsx'.format(dia_indicador.month, dia_indicador.day)
faturamento_loja_dia.to_excel(r'Backup Arquivos Lojas\{}'.format(nome_arquivo))


# ### Passo 6 - Enviar e-mail para diretoria

# In[10]:


outlook = win32.Dispatch('outlook.application')

mail = outlook.CreateItem(0)
mail.To = emails.loc[emails['Loja']=='Diretoria', 'E-mail'].values[0]
mail.Subject = f'Ranking Dia {dia_indicador.day}/{dia_indicador.month}'
mail.Body = f'''
Bom dia, prezados.
Seguem as melhores e piores lojas em faturamento anual e diário.

Melhor Loja do dia em faturamento: Loja {faturamento_loja_dia.index[0]} com faturamento R${faturamento_loja_dia.iloc[0,0]:.2f}
Pior Loja do dia em faturamento: Loja {faturamento_loja_dia.index[-1]} com faturamento R${faturamento_loja_dia.iloc[-1,0]:.2f}
Melhor Loja do ano em faturamento: Loja {faturamento_loja_ano.index[0]} com faturamento R${faturamento_loja_ano.iloc[0,0]:.2f}
Pior Loja do ano em faturamento: Loja {faturamento_loja_ano.index[-1]} com faturamento R${faturamento_loja_ano.iloc[-1,0]:.2f}

Ainda, seguem anexos os arquivos contendo os rankings anual e diário, de todas as nossas lojas!
Atenciosamente, 
Rodrigo Bonfim.
'''

#Anexos
attachment = pathlib.Path.cwd()/caminho_backup/f'{dia_indicador.month}_{dia_indicador.day}_Ranking Anual.xlsx'
mail.Attachments.Add(str(attachment))
attachment = pathlib.Path.cwd()/caminho_backup/f'{dia_indicador.month}_{dia_indicador.day}_Ranking Dia.xlsx'
mail.Attachments.Add(str(attachment))

mail.Send()
print('E-mail da Diretoria enviado')

