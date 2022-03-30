import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
import xlwings as xw
import pandas as pd
import os
import functions as fn

def main():
    open_ell_copy_data()

def get_file_name():
    home_path = str(Path.home())
    
    service_dir = r'\BOLLORE\XPF - Documents\SERVICES'

    paths_joined = os.path.join(home_path + service_dir)

    root = tk.Tk()
    root.lift()
    root.withdraw()

    filename =  filedialog.askopenfilename(initialdir = paths_joined, title = "Select file", filetypes=[("Excel files",".xls .xlsx")])
    root.quit()

    if filename == "":
        exit()
    
    return filename

def open_ell_copy_data():

    ell = get_file_name()

    with xw.App(visible=False) as app:
        wb = app.books.open(ell)
        sheet = wb.sheets('Cargo Detail')
        vessel = sheet.range('A2').value
        voy = sheet.range('B2').value
        leg = sheet.range('C2').value
        
        last_row = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row
        rng_cargo_detail = sheet.range('A5:BF' + str(last_row))
        df = sheet.range(rng_cargo_detail).options(pd.DataFrame, index=False, header=True).value
        df = pd.DataFrame(df).copy()
        wb.close()


    df = df.rename(columns={'ISO Container Type':'ISO TYPE'})
    df.loc[:, 'teus'] = fn.get_TEUs(df)
    df.loc[:, 'mt_or_not'] = df['Commodity'].str.contains('MT')
    df.loc[df['mt_or_not'] == True, 'mt_teus'] = df['teus']
    df.loc[df['mt_or_not'] == False, 'la_teus'] = df['teus']
    df.loc[df['mt_or_not'] == True, 'mt_weight'] = df['Weight in MT']
    df.loc[df['mt_or_not'] == False, 'la_weight'] = df['Weight in MT']
    wb_recap = xw.Book.caller()
    ws_cargo_detail = wb_recap.sheets('Cargo Detail')

    lower_row = ws_cargo_detail.range('A' + str(ws_cargo_detail.cells.last_cell.row)).end('up').row + 1

    if ws_cargo_detail.range('A2').value is None:
        ws_cargo_detail.range('A2').value = vessel
        ws_cargo_detail.range('B2').value = voy
        ws_cargo_detail.range('C2').value = leg

    else:
        if ws_cargo_detail.range('A2').value != vessel:
            messagebox.showwarning(
                title="ELL-info",
                message="VESSEL i ELL stämmer inte överens med föregående fil.")
            exit()

        if ws_cargo_detail.range('B2').value != voy:
            messagebox.showwarning(
                title="ELL-info",
                message="VOYAGE i ELL stämmer inte överens med föregående fil.")
            exit()


    ws_cargo_detail.range('A' + str(lower_row)).options(pd.DataFrame, index=False, header=False).value = df.copy()

    return

if __name__ == '__main__':
    pass