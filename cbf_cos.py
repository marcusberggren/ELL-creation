import pandas as pd
from pandas.api.types import CategoricalDtype
import xlwings as xw
import os as os
from pathlib import Path

def main():

    wb = xw.Book.caller()
    sheet = wb.sheets('Info')

    vessel = sheet.range('B1').value
    alt_voy = sheet.range('I1').value
    pol = sheet.range('N1').value
    date = str(sheet.range('K1').value)

    cell_range = sheet.range('A3').expand() #dynamisk range
    df = sheet.range(cell_range).options(pd.DataFrame, index=False, header=True).value #dynamisk range

    df = df[['TERMINAL', 'ISO TYPE', 'LOAD STATUS', 'NET WEIGHT', 'VGM']]

    #När NET WEIGHT är mindre än 100 men större än 0 så multiplicera med 1000
    df.loc[(df['NET WEIGHT'] < 100) & (df['NET WEIGHT'] != 0), 'NET WEIGHT'] *= 1000

    #Skapar ny kolumn med maxvärde av NET WEIGHT & VGM
    df['MAX WEIGHT'] = df[['NET WEIGHT', 'VGM']].max(axis=1) // 1000  

    df = df[['TERMINAL', 'ISO TYPE', 'LOAD STATUS', 'MAX WEIGHT']]

    df.loc[df['LOAD STATUS'] != 'MT', 'LOAD STATUS'] = 'LA'

    df.insert(4, 'WEIGHT_TYPE', "")

    df.loc[(df['LOAD STATUS'] == 'LA') & (df['ISO TYPE'].str[0] == '2') & (df['MAX WEIGHT'] >= 20), 'WEIGHT_TYPE'] = 'VH20'
    df.loc[(df['LOAD STATUS'] == 'LA') & (df['ISO TYPE'].str[0] == '2') & (df['MAX WEIGHT'] >= 15) & (df['MAX WEIGHT'] < 20), 'WEIGHT_TYPE'] = 'H20'
    df.loc[(df['LOAD STATUS'] == 'LA') & (df['ISO TYPE'].str[0] == '2') & (df['MAX WEIGHT'] >= 10) & (df['MAX WEIGHT'] < 15), 'WEIGHT_TYPE'] = 'M20'
    df.loc[(df['LOAD STATUS'] == 'LA') & (df['ISO TYPE'].str[0] == '2') & (df['MAX WEIGHT'] < 10), 'WEIGHT_TYPE'] = 'L20'

    df.loc[(df['LOAD STATUS'] == 'LA') & (df['ISO TYPE'].str[:2] == '42') & (df['MAX WEIGHT'] >= 25), 'WEIGHT_TYPE'] = 'VH42'
    df.loc[(df['LOAD STATUS'] == 'LA') & (df['ISO TYPE'].str[:2] == '42') & (df['MAX WEIGHT'] >= 20) & (df['MAX WEIGHT'] < 25), 'WEIGHT_TYPE'] = 'H42'
    df.loc[(df['LOAD STATUS'] == 'LA') & (df['ISO TYPE'].str[:2] == '42') & (df['MAX WEIGHT'] >= 15) & (df['MAX WEIGHT'] < 20), 'WEIGHT_TYPE'] = 'M42'
    df.loc[(df['LOAD STATUS'] == 'LA') & (df['ISO TYPE'].str[:2] == '42') & (df['MAX WEIGHT'] < 15), 'WEIGHT_TYPE'] = 'L42'

    df.loc[(df['LOAD STATUS'] == 'LA') & (df['ISO TYPE'].str[:2] == '45') & (df['MAX WEIGHT'] >= 25), 'WEIGHT_TYPE'] = 'VH45'
    df.loc[(df['LOAD STATUS'] == 'LA') & (df['ISO TYPE'].str[:2] == '45') & (df['MAX WEIGHT'] >= 20) & (df['MAX WEIGHT'] < 25), 'WEIGHT_TYPE'] = 'H45'
    df.loc[(df['LOAD STATUS'] == 'LA') & (df['ISO TYPE'].str[:2] == '45') & (df['MAX WEIGHT'] >= 15) & (df['MAX WEIGHT'] < 20), 'WEIGHT_TYPE'] = 'M45'
    df.loc[(df['LOAD STATUS'] == 'LA') & (df['ISO TYPE'].str[:2] == '45') & (df['MAX WEIGHT'] < 15), 'WEIGHT_TYPE'] = 'L45'

    df.loc[(df['LOAD STATUS'] == 'MT') & (df['ISO TYPE'].str[0] == '2'), 'WEIGHT_TYPE'] = 'MT20'
    df.loc[(df['LOAD STATUS'] == 'MT') & (df['ISO TYPE'].str[:2] == '42'), 'WEIGHT_TYPE'] = 'MT42'
    df.loc[(df['LOAD STATUS'] == 'MT') & (df['ISO TYPE'].str[:2] == '45'), 'WEIGHT_TYPE'] = 'MT45'

    lista_iso_type = ['VH20', 'VH42', 'VH45', 'H20', 'H42', 'H45', 'M20', 'M42', 'M45', 'L20', 'L42', 'L45', 'MT20', 'MT42', 'MT45']

    cat_size_order =CategoricalDtype(lista_iso_type, ordered=True)

    df['WEIGHT_TYPE'] = df['WEIGHT_TYPE'].astype(cat_size_order)

    df.sort_values(['WEIGHT_TYPE', 'TERMINAL'], ascending=(True, False))

    df = df.groupby(['TERMINAL', 'WEIGHT_TYPE']).size().iteritems()        #groupby och size för att skapa hanterlig strukturerad data. Kan iterera över datan nedan mha iteritems

    

    dict1 = {}

    for (terminal, iso_type), antal in df:              #tar fram nestlad info (terminal, iso_type) och antal från ovan
        if terminal not in dict1:
            dict1[terminal] = {}
        dict1[terminal].update({iso_type:antal})

    terminal = ""
    index = 0

    lista_weight_type = []
    create_nested_list = []
    lista_terminal = []
    lista_unika_terminaler = []

    for terminal in dict1:
        lista_unika_terminaler.append(terminal)

        for index, weight_type in enumerate(dict1[terminal].values()):

            if index % 3 == 2:
                lista_weight_type.append(weight_type)
                create_nested_list.append(lista_weight_type)
                lista_weight_type = []
            else:
                lista_weight_type.append(weight_type)
        lista_terminal.append(create_nested_list)
        create_nested_list = []

    df_new = pd.DataFrame(lista_terminal)

    def get_terminal(lista_terminaler):
        i = 0
        for i, terminal in enumerate(lista_terminaler):

            if terminal == "NLEDE":
                lista_terminaler[i] = "RTM-DDE"
            elif terminal == "NLEMX":
                lista_terminaler[i] = "RTM-EMX"
            elif terminal == "NLRWG":
                lista_terminaler[i] = "RTM-RWG"
            elif terminal == "DECTB":
                lista_terminaler[i] = "HAM-CTB"
            elif terminal == "DECTA":
                lista_terminaler[i] = "HAM-CTA"
            elif terminal == "DETCT":
                lista_terminaler[i] = "HAM-CTT"

        return lista_terminaler

    with xw.App(visible=False) as app:
        home = Path.home()
        whole_path = home / 'BOLLORE\XPF - Documents\MAINTENANCE\MLO_COS_CBF_SEGOT.xls'
        wb = app.books.open(whole_path)
        
        wb_caller = xw.Book.caller()
        wb_caller_name = wb_caller.fullname
        folder_path = os.path.split(wb_caller_name)[0]
        filename = 'CBF_'+ vessel + '_' + str(alt_voy) +'_' + pol + '.xlsx'
        dir_path = os.path.join(folder_path, filename)
        
        ws = wb.sheets['CBF TTL']
        ws.range('B3').value = vessel
        ws.range('B4').value = alt_voy
        ws.range('I3').value = pol
        ws.range('I4').value = date[:10]

        terminal, cell, i = "", 8, 0
        for i, terminal in enumerate(get_terminal(lista_unika_terminaler)):
            
            ws.range('A' + str(cell + 1)).options(index=False, header=False).value = terminal
            
            ws.range('C' + str(cell)).options(index=False, header=False).value = df_new[0][i]
            ws.range('C' + str(cell + 1)).options(index=False, header=False).value = df_new[1][i]
            ws.range('C' + str(cell + 2)).options(index=False, header=False).value = df_new[2][i]
            ws.range('C' + str(cell + 3)).options(index=False, header=False).value = df_new[3][i]
            ws.range('C' + str(cell + 4)).options(index=False, header=False).value = df_new[4][i]
            cell += 7

        wb.save(dir_path)
        wb.close()