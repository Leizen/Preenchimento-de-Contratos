#!/usr/bin/env python
# coding: utf-8

# In[1]:


from mailmerge import MailMerge
import pandas as pd, openpyxl

#Template/Documento docx que será preenchido.
template = "Contrato.docx"

#Criando o documento através do MailMerge
document = MailMerge(template)

#Printando todos os campos preenchiveis no documento
print(document.get_merge_fields())


# In[5]:


#Lendo o arquivo Excel onde estão as informações para preencher o docx e criando um DataFrame
#Utilizando o openpyxl para leitura e setando todos os campos como String
df = pd.read_excel("Dados.xlsx", engine='openpyxl').astype(str)

#Trocando os campos vazios "nan" por um campo em branco
df.replace({"nan":" "}, inplace=True)


# In[8]:


#Percorrendo linha a linha do DataFrame
for row in df.itertuples():
    
    #Para cada linha, ele irá criar um novo documento    
    document = MailMerge(template)

    #Para cada campo do documento, ele vai atribuir o valor da sua respectiva coluna
    document.merge(
        CEP = row.CEP,
        NomeEmpresa = row.EMPRESA,
        Complemento = row.COMPLEMENTO,
        Bairro = row.BAIRRO,
        Estado = row.ESTADO,
        Representante = row.REPRESENTANTE,
        CNPJ = row.CNPJ,
        Endereco = row.ENDERECO
    )
    
    #Escreve o documento com os valores preenchidos com o nome de CONTRATO e NOME DA EMPRESA e fecha o documento
    document.write("CONTRATO "+row.EMPRESA+".docx")
    document.close()

