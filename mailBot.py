#!/usr/bin/env python
# coding: utf-8

# ## Automatizando envio de e-mails para um boletim informativo.

# ### Descrição

# O código abaixo auxilia em uma rotina de envio de e-mails para boletim informativo da empresa X.
# 
# Essa automação usurá o pacote pandas e win32com (Integração Python com Outlook), é possível utilizar o pacote yagmail para integração com o gmail.
# 
# Pegaremos o nome dos clientes e seus respectivos e-mails atráves de um arquivo excel*.
# 
# *Excel -> Utilizando de uma ETL é possível montar uma tabela com nome e e-mail dos clientes atráves do banco de dados da empresa.*

# In[1]:


#Importando pandas
import pandas as pd

#Importando win32com
import win32com.client as win32
outlook = win32.Dispatch('outlook.application')

#Importando e lendo o arquivo excel
clientesDataframe = pd.read_excel('Clientes.xlsx')

#Imprimindo o dataframeClientes
display(clientesDataframe)

#Percorrendo nosso arquivo, recuperando os nomes e e-mails
for i, email in enumerate(clientesDataframe['E-mail']):
    nome = clientesDataframe.loc[i, 'Nome']
    
    mail = outlook.CreateItem(0)
    mail.To = email
    mail.Subject = 'Boletim Informativo - X'
    mail.HTMLBody = '''
    <p>Olá, <b>{}</b></p><h4>Este é nosso boletim informativo mensal.</h4>
    '''.format(nome)
    
    mail.Send()
    

