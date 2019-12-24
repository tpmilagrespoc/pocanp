# -*- coding: utf-8 -*-
"""
Created on Mon Dec 23 17:50:09 2019

@author: tpmilagres
"""


import pandas as pd
import numpy as np
import urllib
import win32com.client
import psycopg2
import os
import time
from datetime import datetime

url = 'http://www.anp.gov.br/images/DADOS_ESTATISTICOS/Vendas_de_Combustiveis/Vendas_de_Combustiveis_m3.xls'
vba_dir = os.path.dirname(os.path.realpath(__file__))+"\\vba\\vba.xlsm"


def bulk_load(df, table):
    ## BulkLoad
    con = psycopg2.connect(host='34.95.173.145', database = 'poc_anp',
                          user = 'user', password='pwd')
    cur = con.cursor()

    csv = df.to_csv('file.csv',index=False, sep=',',header=False,encoding='cp1252')
    
    cur.execute('truncate table '+ table)
    con.commit()
    
    with open ('file.csv','r') as f:
        cur.copy_from(f, table , sep=',')

    con.commit()



#Recupera arquivo
file_dir, head = urllib.request.urlretrieve(url)

#Vba para expandir PivotCache
data=win32com.client.Dispatch("Excel.Application")
vba = win32com.client.Dispatch("Excel.Application")


vba.Workbooks.open(vba_dir)
data.Workbooks.open(file_dir)

data.Application.Run("vba.xlsm!extract_cache.extract")

#Acesso a PivotCache
df_data = pd.read_excel(file_dir,sheet_name='Sheet2', header = 0)

#Group By
df_gp = df_data.fillna(0).groupby(['COMBUSTÍVEL', 'ANO', 'REGIÃO', 'ESTADO', 'UNIDADE', 'Jan', 'Fev',
       'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'], as_index=False).sum()



#Unpivot
df_pvt = pd.melt(df_gp,
                       value_vars  = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'],
                      id_vars= ['COMBUSTÍVEL', 'ANO', 'REGIÃO', 'ESTADO', 'UNIDADE']
                       
    )
  
#Depara Ano
df_replace = df_pvt.replace({'variable':{'Jan':'1','Fev':'2',
       'Mar':'3', 'Abr':'4', 'Mai':'5', 'Jun':'6', 'Jul':'7', 'Ago':'8', 'Set':'9', 'Out':'10', 'Nov':'11', 'Dez':'12'}})

#Ajustar nome das colunas
df_replace.columns=['produto', 'ano', 'regiao', 'estado', 'unidade','mes','vol_demanda_m3']

#inclui timestamp captura
df_replace['timestamp_captura']=datetime.fromtimestamp(time.time())

#remover coluna regiao para insert
df_replace = df_replace.drop(columns=['regiao'])


#Insert Por Bulk Load (copy)

bulk_load(df=df_replace,table='stg.venda_produto')


