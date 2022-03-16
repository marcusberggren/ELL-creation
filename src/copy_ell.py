import pandas as pd
import xlwings as xw
import os as os
import tkinter as tk
from tkinter import filedialog, messagebox
import functions as fn

def main():
    caller_wb, caller_sheet = set_variables()
    last_row = get_last_row(caller_sheet)
    chosen_file, weight_status = open_ell(caller_wb)
    dataframe = create_dataframe(chosen_file, weight_status)
    copy_data(dataframe, caller_sheet, last_row)

def set_variables():
    wb_caller = xw.Book.caller()
    folder_path = os.path.split(wb_caller.fullname)[0]
    caller_sheet = wb_caller.sheets['INFO']

    return folder_path, caller_sheet

def get_last_row(sheet):
    lower_row = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row

    for row in range(lower_row, 1, -1):
        if sheet.range('A' + str(row)).value == None:
            next
        elif row == 5:
            return 5
        else:
            return row

def open_ell(path):
    root = tk.Tk()
    root.lift()
    root.withdraw()

    choose_file =  filedialog.askopenfilename(initialdir = path, title = "Select file", filetypes=[("Excel files",".xls .xlsx")])

    if choose_file != "":
        message_box = messagebox.askyesno('ELL-fil', 'Är vikterna preliminära i filen?')
        if message_box:
            svar = "yes"
        else:
            svar = "no"
    else:
        svar = "yes"

    root.quit()
    return choose_file, svar

def create_dataframe(path, pre_vgm_weights):
    df_cargo_detail = pd.read_excel(path, sheet_name='Cargo Detail', header=4)
    df_manifest = pd.read_excel(path, sheet_name='Manifest', header=0)

    df_cargo_detail.columns = map(str.upper, df_cargo_detail.columns)
    df_manifest.columns = map(str.upper, df_manifest.columns)

    df_cargo_detail = df_cargo_detail[['POD', 'POD TERMINAL','BOOKING REFERENCE', 'MLO PO', 'MLO', 'MOTHER VESSEL', 'MOTHER VOYAGE',
    'F.DESTINATION','CARGO TYPE', 'COMMODITY', 'CONTAINER NO', 'WEIGHT IN MT', 'TEMPMAX', 'IMCO', 'UN', 'VGM WEIGHT IN MT']].copy()

    df_manifest = df_manifest[['GOODS DESC', 'NO OF PACKAGES', 'NET WEIGHT IN KILOS']].copy()

    df_final = pd.DataFrame(columns=['BOOKING NUMBER', 'MLO', 'POL', 'TOL', 'CONTAINER', 'ISO TYPE',
    'NET WEIGHT', 'POD STATUS', 'LOAD STATUS', 'VGM', 'OOG', 'REMARK', 'IMDG', 'UNNR', 'CHEM REF', 'MRN', 'TEMP',
    'PO NUMBER', 'CUSTOMS STATUS', 'PACKAGES', 'GOODS DESCRIPTION', 'OCEAN VESSEL', 'VOYAGE', 'ETA', 'FINAL POD'])

    df_final.loc[:, 'BOOKING NUMBER'] = df_cargo_detail['BOOKING REFERENCE']
    df_final.loc[:, 'MLO'] = df_cargo_detail['MLO']
    df_final.loc[:, 'POL'] = df_cargo_detail['POD']
    df_final.loc[:, 'TOL'] = df_cargo_detail['POD TERMINAL']
    df_final.loc[:, 'CONTAINER'] = df_cargo_detail['CONTAINER NO']
    df_final.loc[:, 'ISO TYPE'] = df_cargo_detail['CARGO TYPE']
    
    if pre_vgm_weights == "yes":
        df_final.loc[:, 'NET WEIGHT'] = df_cargo_detail[['WEIGHT IN MT', 'VGM WEIGHT IN MT']].max(axis=1)
        df_final.loc[:, 'VGM'] = ''
    elif pre_vgm_weights == "no":
        df_final.loc[:, 'NET WEIGHT'] = df_manifest['NET WEIGHT IN KILOS']
        df_final.loc[:, 'VGM'] = df_cargo_detail[['WEIGHT IN MT', 'VGM WEIGHT IN MT']].max(axis=1) * 1000

    df_final.loc[:, 'POD STATUS'] = "T"
    df_final.loc[:, 'LOAD STATUS'] = df_cargo_detail['COMMODITY']
    df_final.loc[:, 'OOG'] = ""
    df_final.loc[:, 'REMARK'] = ""
    df_final.loc[:, 'IMDG'] = df_cargo_detail['IMCO']
    df_final.loc[:, 'UNNR'] = df_cargo_detail['UN']
    df_final.loc[:, 'CHEM REF'] = ""
    df_final.loc[:, 'MRN'] = ""
    df_final.loc[:, 'TEMP'] = df_cargo_detail['TEMPMAX']
    df_final.loc[:, 'PO NUMBER'] = df_cargo_detail['MLO PO']
    df_final.loc[:, 'CUSTOMS STATUS'] = ""
    df_final.loc[:, 'PACKAGES'] = ""
    df_final.loc[:, 'GOODS DESCRIPTION'] = df_manifest['GOODS DESC']
    df_final.loc[:, 'OCEAN VESSEL'] = df_cargo_detail['MOTHER VESSEL']
    df_final.loc[:, 'ETA'] = ""
    df_final.loc[:, 'VOYAGE'] = df_cargo_detail['MOTHER VOYAGE']
    df_final.loc[:, 'FINAL POD'] = df_cargo_detail['F.DESTINATION']
    if df_final['POL'].iloc[-1] == 'END':
        df_final = df_final[:-1]
    return df_final

def copy_data(dataframe, data_to_sheet, row_number):
    row = str(row_number + 1)
    data_to_sheet['A' + row].options(pd.DataFrame, header=None, index=False).value = dataframe
    return

if __name__ == '__main__':
    file_path = fn.get_mock_caller('0109_Bokningsblad.xlsb')
    xw.Book(file_path).set_mock_caller()
    main()
