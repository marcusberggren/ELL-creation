# Functions to be imported
import xlwings as xw
import pandas as pd
import numpy as np
from pathlib import Path
import re

def get_path(text_input: str):
    home = str(Path.home())
    func_path = get_path.__code__.co_filename
    trimmed_path = re.sub(r"\w+\\\w+\.py$", "", func_path)    #Tar bort sista två orden + .py i path
    file_path = trimmed_path + 'data\stored-data-paths.csv'
    df_csv = pd.read_csv(file_path, sep=';', index_col=0, skipinitialspace=True)
    new_dict = df_csv.to_dict()['PATH']
    return_path = home + new_dict[text_input]
    return return_path

def get_csv_data(path_name: str):
    file_path = get_path(path_name)
    file = pd.read_csv(file_path, sep=';', header=0, index_col=None, skipinitialspace=True)
    df = pd.DataFrame(file, index=None)
    return df

def get_caller_df():
    wb = xw.Book.caller()
    sheet = wb.sheets('INFO')
    data_table = sheet.range('A4').expand()
    df = sheet.range(data_table).options(pd.DataFrame, index=False, header=True).value
    df = regex_no_extra_whitespace(df).copy()
    
    get_caller_df.vessel = sheet.range('A2').value
    get_caller_df.voyage = str(sheet.range('B2').value)
    get_caller_df.pol = sheet.range('D2').value
    return df

def get_mock_caller(excel_file: str):
    func_path = get_path.__code__.co_filename
    trimmed_path = re.sub(r"\w+\\\w+\.py$", "", func_path)    #Tar bort sista två orden + .py i path
    file_path = trimmed_path + 'data\\' + excel_file
    return file_path

def regex_no_extra_whitespace(df: pd.DataFrame):
    df = df.replace(r"^\s+|\s+$", "", regex=True).copy()
    return df

def container_check(container_no: str):
    var_dict = {
        "A":10, "B":12, "C":13, "D":14, "E":15, "F":16, "G":17,
        "H":18, "I":19, "J":20, "K":21, "L":23, "M":24, "N":25,
        "O":26, "P":27, "Q":28, "R":29, "S":30, "T":31, "U":32,
        "V":34, "W":35, "X":36, "Y":37, "Z":38
        }

    value_multiply, summa, = 0, 0
    len_cont = len(container_no)

    if container_no[:3] == "DUM":
        return False
    elif len_cont != 11:
        return False
    else:
        for num, character in enumerate(container_no):
            if num == 0:
                value_multiply = 1
            elif num == 10:
                continue
            else:
                value_multiply *= 2

            if re.search('[a-zA-z]', character):
                summa += int(var_dict.get(character)) * value_multiply
            elif re.search('[0-9]', character):
                summa += int(character) * value_multiply

        sum_changed = int(summa/11) * 11
 
        if summa - sum_changed == 10 and int(container_no[len_cont-1]) == 0:
            return True
        elif summa - sum_changed == int(container_no[len_cont-1]):
            return True
        else:
            return False

def terminal_check(df: pd.DataFrame):
    df_csv = get_csv_data('terminal').copy()

    df_csv['CONCAT'] = df_csv['PORT'] + df_csv['TERMINAL']
    df['CONCAT'] = df['POL'] + df['TOL']

    df.loc[df['CONCAT'].isin(df_csv['CONCAT']), 'TERMINAL_CHECK'] = True
    df.loc[np.logical_not(df['CONCAT'].isin(df_csv['CONCAT'])), 'TERMINAL_CHECK'] = False
    return df['TERMINAL_CHECK']

def MLO_check(df: pd.DataFrame):
    df_csv = get_csv_data('mlo').copy()
    df.loc[df['MLO'].isin(df_csv['MLO']), 'MLO_CHECK'] = True
    df.loc[np.logical_not(df['MLO'].isin(df_csv['MLO'])), 'MLO_CHECK'] = False
    return df['MLO_CHECK']

def cargo_type_check(df: pd.DataFrame):
    df_csv = get_csv_data('cargo_type').copy()
    df['ISO TYPE'] = df['ISO TYPE'].astype(str)

    df['CONCAT'] = df['ISO TYPE'] +  df['LOAD STATUS']
    df.loc[df['CONCAT'].isin(df_csv['ISO STATUS']), 'CARGO_TYPE_CHECK'] = True
    df.loc[np.logical_not(df['CONCAT'].isin(df_csv['ISO STATUS'])), 'CARGO_TYPE_CHECK'] = False
    return df['CARGO_TYPE_CHECK']

def load_status_check(df: pd.DataFrame):
    df_csv = get_csv_data('load_status').copy()
    df.loc[df['LOAD STATUS'].isin(df_csv['LOAD STATUS']), 'LOAD_STATUS_CHECK'] = True
    df.loc[np.logical_not(df['LOAD STATUS'].isin(df_csv['LOAD STATUS'])), 'LOAD_STATUS_CHECK'] = False
    df.loc[df['LOAD STATUS'].str.contains("MT"), 'LOAD_STATUS_CHECK'] = "MT"
    return df['LOAD_STATUS_CHECK']

def reefer_check(df: pd.DataFrame):
    df.loc[:, 'TEMP_CHECK'] = True
    df.loc[(df['ISO TYPE'].str.contains("R1")) & (df['TEMP'].isnull()), 'TEMP_CHECK'] = False
    df.loc[(df['LOAD STATUS'].str.contains("RF")) & (df['TEMP'].isnull()), 'TEMP_CHECK'] = False
    df.loc[(np.logical_not(df['ISO TYPE'].str.contains("R1"))) &    #np.logical_not to reverse the boolean
        (np.logical_not(df['LOAD STATUS'].str.contains("RF"))) &
        (df['TEMP'].notnull()), 'TEMP_CHECK'] = False
    df.loc[(df['ISO TYPE'].str.contains("R1")) & (df['TEMP'].notnull()), 'TEMP_CHECK'] = True
    return df['TEMP_CHECK']

def dg_check(df: pd.DataFrame):
    df.loc[:, 'DG_CHECK'] = True
    df.loc[(df['LOAD STATUS'].str.contains('DG')) & (df['IMDG'].isnull()), 'DG_CHECK'] = False
    df.loc[(df['LOAD STATUS'].str.contains('DG')) & (df['UNNR'].isnull()), 'DG_CHECK'] = False
    df.loc[(np.logical_not(df['LOAD STATUS'].str.contains('DG'))) & (df['IMDG'].notnull()), 'DG_CHECK'] = False
    df.loc[(np.logical_not(df['LOAD STATUS'].str.contains('DG'))) & (df['UNNR'].notnull()), 'DG_CHECK'] = False
    df.loc[(df['IMDG'].notnull()) & (df['UNNR'].isnull()), 'DG_CHECK'] = False
    df.loc[(np.logical_not(df['IMDG'].notnull())) & (df['UNNR'].notnull()), 'DG_CHECK'] = False
    df.loc[df['LOAD STATUS'].str.contains('DG') & (df['IMDG'].notnull()) & (df['UNNR'].notnull()), 'DG_CHECK'] = True
    return df['DG_CHECK']

def fpod_check(df: pd.DataFrame):
    df_csv = get_csv_data('ports').copy()
    df.loc[df['FINAL POD'].isin(df_csv['CODE']), 'FPOD_CHECK'] = True
    df.loc[np.logical_not(df['FINAL POD'].isin(df_csv['CODE'])), 'FPOD_CHECK'] = False
    return df['FPOD_CHECK']

def vessel_check(df: pd.DataFrame):
    df_csv = get_csv_data('ocean_vessel').copy()
    df.loc[df['OCEAN VESSEL'].isin(df_csv['OCEAN VESSEL']), 'VESSEL_CHECK'] = True
    df.loc[np.logical_not(df['OCEAN VESSEL'].isin(df_csv['OCEAN VESSEL'])), 'VESSEL_CHECK'] = False
    return df['VESSEL_CHECK']

def customs_check(df: pd.DataFrame):
    df_csv = get_csv_data('eu')

    
    df.loc[(df['FINAL POD'].str[:2].isin(df_csv['EU COUNTRIES'])) & (df['POL'] == "NLRTM"), 'CUSTOMS_CHECK'] = "X"                  #NLRTM
    df.loc[(np.logical_not(df['CUSTOMS STATUS'].isin(df_csv['EU COUNTRIES']))) & (df['POL'] == "NLRTM"), 'CUSTOMS_CHECK'] = "N"     #NLRTM
    df.loc[(np.logical_not(df['CUSTOMS STATUS'].isin(df_csv['EU COUNTRIES']))) & (df['POL'] == "BEANR"), 'CUSTOMS_CHECK'] = "N"     #BEANR
    df.loc[(np.logical_not(df['CUSTOMS STATUS'].isin(df_csv['EU COUNTRIES']))) & (df['POL'] == "DEHAM"), 'CUSTOMS_CHECK'] = "T1"    #DEHAM
    df.loc[(np.logical_not(df['CUSTOMS STATUS'].isin(df_csv['EU COUNTRIES']))) & (df['POL'] == "DEBRV"), 'CUSTOMS_CHECK'] = "T1"    #DEBRV
    df.loc[df['FINAL POD'].str[:2].isin(df_csv['EU COUNTRIES']), 'CUSTOMS_CHECK'] = "C"                                             #EU country
    df.loc[df['LOAD STATUS'].str.contains("MT"), 'CUSTOMS_CHECK'] = "C"                                                             #Empty
    return df['CUSTOMS_CHECK']


def get_max_weight(df: pd.DataFrame):
    df['VGM'] = df['VGM'].fillna(0)
    df.loc[(df['NET WEIGHT'] >= 100) & (df['VGM'] == 0), 'WEIGHT+TARE'] = df[['NET WEIGHT', 'TARE']].sum(axis=1)
    df.loc[(df['NET WEIGHT'] < 100) & (df['NET WEIGHT'] != 0), 'WEIGHT+TARE'] = df['NET WEIGHT'] * 1000
    df.loc[df['VGM'] > 0, 'WEIGHT+TARE'] = df[['NET WEIGHT', 'VGM']].max(axis=1)
    return df['WEIGHT+TARE']

def get_TEUs(df: pd.DataFrame):
    conditions_teu = [
            (df['ISO TYPE'].str[:1] == "2"),
            (df['ISO TYPE'].str[:1] == "3"),
            (df['ISO TYPE'].str[:1] == "4"),
            (df['ISO TYPE'].str[:1] == "L")
        ]
    values_teu = [1, 2, 2, 2]
    result = np.select(conditions_teu, values_teu)
    return result

def get_tare(df: pd.DataFrame):
    conditions_tare = [
            (df['ISO TYPE'].str[:1] == "2"),
            (df['ISO TYPE'].str[:1] == "3"),
            (df['ISO TYPE'].str[:1] == "4"),
            (df['ISO TYPE'].str[:1] == "L")
        ]
    values_tare = [2200, 3200, 4000, 4000]
    result = np.select(conditions_tare, values_tare)
    return result

def get_template_type(df: pd.DataFrame, template: list):
    file_path = get_path(template[0])
    df_csv = pd.read_csv(file_path, sep=';', index_col=0, skipinitialspace=True)
    new_dict = df_csv.to_dict()[template[1]]
    df = df[template[2]].replace(new_dict).copy()
    return df