from email import header
import requests
import pandas as pd
import xlwings as xw
import time
import datetime
import os
import ctypes
import numpy as np

def tratar_id(id_pacote_func):
    concatenado = ''
    if id_pacote_func == '':
        pass
    else:
        for strr in id_pacote_func:
            try:
                int(strr)
                concatenado = concatenado + strr
            except:
                pass
        try:
            return float(concatenado)
        except:
            return float('nan')

def loop_atualizacao():

    # baixar_planilha()

    tabela_inv = pd.read_excel('https://docs.google.com/spreadsheets/d/1JcQH_Df-_jnFMGJoFXVk4yLr43mQTo5elD2VObd2iOw/export?format=xlsx')
    tabela_vuc = pd.read_excel('https://docs.google.com/spreadsheets/d/1GIlraHslu0FahZK6-KLEE4GY7odiIMHfg9TKTv3z0FI/export?format=xlsx')
    tabela_dev = pd.read_excel('https://docs.google.com/spreadsheets/d/1jK_jw6vhLc03MLJyfdZ2nNqiTYal8zyDdngM_kOQ1ZM/export?format=xlsx')
    expedicao_vuc = pd.read_excel('https://docs.google.com/spreadsheets/d/1pPOhGrHQBm8Rac09b9yhrCULD0nA5Czo8Wvsp0Lh47k/export?format=xlsx')
    expedicao_vuc = expedicao_vuc[['Carimbo de data/hora',	'Transportadora',	'Shipments']]
    expedicao_vuc['Área'] = expedicao_vuc['Transportadora']
    expedicao_vuc = expedicao_vuc[['Carimbo de data/hora',	'Área',	'Shipments']]
    tabela_vuc['Responsável pela bipagem'] = np.nan
    tabela_dev['Responsável pela bipagem'] = np.nan
    expedicao_vuc['Responsável pela bipagem'] = np.nan
    tabela_inv = pd.concat([tabela_inv,tabela_vuc,expedicao_vuc,tabela_dev])
    tabela_inv['Shipments'] = tabela_inv['Shipments'].astype('str')
    tabela_inv['Área'] = tabela_inv['Área'].astype('str')
    tabela_pacotes_inventariados = pd.DataFrame(columns=['Shipments','Área','Carimbo de data/hora','Responsável pela bipagem'])
    
    for linha in range(len(tabela_inv)):

        area_i = tabela_inv[['Área']].values[linha][0]
        data_e_hora = tabela_inv[['Carimbo de data/hora']].values[linha][0]
        responsavelbipagem = tabela_inv[['Responsável pela bipagem']].values[linha][0]
        shipments_i = pd.Series(str(tabela_inv[['Shipments']].values[linha][0]).split('\n'))
        shipments_j = []

        for linha_shipments in shipments_i:
            try:
                int(str(linha_shipments)[0])
                shipments_j.append(linha_shipments)
            except:
                shipments_j.append(tratar_id(linha_shipments))
    
        tabela_xpt = pd.DataFrame(shipments_j, columns=['Shipments'])
        tabela_xpt[['Área']] = area_i
        tabela_xpt[['Carimbo de data/hora']] = data_e_hora
        tabela_xpt[['Responsável pela bipagem']] = responsavelbipagem

        tabela_pacotes_inventariados = pd.concat([tabela_pacotes_inventariados,tabela_xpt])
        
    tabela_pacotes_inventariados = tabela_pacotes_inventariados.loc[~ (tabela_pacotes_inventariados['Shipments'] == '')]
    #tabela_pacotes_inventariados = tabela_pacotes_inventariados.loc[tabela_pacotes_inventariados['Carimbo de data/hora'] >= datetime.datetime(int(time.strftime("%Y")),int(time.strftime("%m")),int(time.strftime("%d")))]
    tabela_pacotes_inventariados['Shipments'] = pd.to_numeric(tabela_pacotes_inventariados['Shipments'], errors='coerce')
    tabela_pacotes_inventariados = tabela_pacotes_inventariados.loc[~ (tabela_pacotes_inventariados['Shipments'].isna())]
    tabela_pacotes_inventariados = tabela_pacotes_inventariados.loc[~ (tabela_pacotes_inventariados['Shipments'] > 50000000000)]
    tabela_pacotes_inventariados['Shipments'] = tabela_pacotes_inventariados['Shipments'].astype('int64')
    # tabela_pacotes_inventariados = tabela_pacotes_inventariados.drop_duplicates(subset=['Shipments','Área'])
    tabela_pacotes_inventariados['Área'] = tabela_pacotes_inventariados['Área'].astype('str')

    return tabela_pacotes_inventariados

app = xw.App(visible = False, add_book = False)

wb = xw.Book(r"Inventário de Pacotes.xlsm")
wb.activate()
ws = wb.sheets["base_forms"]
app.screen_updating = False
ws["a:d"].clear_contents()
ws["A1"].options(index=False).value = loop_atualizacao()
app.screen_updating = True
