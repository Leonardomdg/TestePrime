#importar as libs necessárias:


import pandas as pd
from selenium import webdriver as wd
from time import sleep
from os import getenv

#armazenamento da url e abertura do navegador Chrome

url = "https://protect-de.mimecast.com/s/nQAhCmq14NtRjJZQxHG2Vl3?domain=itdashboard.gov"
driver = wd.Chrome(f'{getenv("localappdata")}\\chromedriver.exe')
driver.maximize_window()
driver.get(url)

#busca por agencias

driver.find_element("xpath",'//*[@id="edit-keywords"]').send_keys('Agency')
sleep(2)
driver.find_element("xpath",'//*[@id="edit-submit"]').click()

#Baixa arquivo resumido
sleep(2)
driver.find_element("xpath",'//*[@id="download-csv-results"]').click()

#Montagem de dataframe

df = pd.read_csv('C:/Users/leona/Downloads/Investment-Results-Details (2).csv',
    sep=',',on_bad_lines='skip')#Abertura e leitura, onbadlines corrige erros
df = pd.DataFrame(df) #garante a transformação pra dataframe

colunas = {} #Dicionário instanciado em vazio
for i in range(len(df.columns)-1): #Pega da primeira coluna até a penúltima coluna
    colunas[df.columns[i]] = df.columns[i+1] #passa o que está na coluna da direita para esquerda
colunas[df.columns[len(df.columns)-1]]=df.columns[0] #esse atualiza a última coluna

df.rename(columns = colunas,inplace=True) #columns recebe as colunas de cima, inplace atualiza o dataframe
df.drop(columns=df.columns[len(df.columns)-1],inplace=True) #Esse apaga a última inútil[0]

df.to_excel('ArchiveCompleted.xlsx')

AgencyVar = list(df['Agency'].unique())

dflistas = []
for agencia in AgencyVar:
    dflistas.append(df[df['Agency']==agencia])

writer = pd.ExcelWriter('ArchiveCompleted.xlsx',engine='xlsxwriter')
df.to_excel(writer,sheet_name='Base Agencias',index=False)
for i,agencia in enumerate(dflistas):
    agencia.to_excel(writer,sheet_name=f'Agencia{i+1}',index=False) #Exporta pra excel pela pip openpyxl
writer.save()
writer.close()