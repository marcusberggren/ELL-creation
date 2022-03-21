import re
import pandas as pd
import tkinter as tk
from tkinter import filedialog
import xlwings as xw

def main():
    copy_data()

def chosen_file():
    root = tk.Tk()
    root.lift()
    root.withdraw()
    
    filename =  filedialog.askopenfilename(initialdir = xw.Book.caller(), title = "Select file", filetypes=[("Txt- & Excel files",".txt .xls .xlsb")])
    root.quit()
    
    if filename == "":
        exit()

    return filename

def get_data():
    new_list = []
    hl_list = []
    booking_no = ""
    all_details = ""
    textfile = chosen_file()
    
    with open(textfile) as fn:
        row_count = 0
        tod_data = ""
        amount = 0
        container = ""

        for row, line in enumerate(fn):
            if len(line) < 2:
                continue
            
            if "PORT-OF-DISCHARGE:" in line:
                tod_data = line[25:-1].replace(" ", "")         # Byter ut mellanslag
                pod_data = line[19:24]

            line = re.sub(r"\A\s{1,8}", booking_no, line)       # Om radens inledning består av 8st mellanslag byt det ut mot bokningsnumret.
            line = re.sub(r"\A.{9}\s{37}", all_details, line)   # Om radens inledning är 9 st X-bokstäver/siffror men efterföljande 37 symboler är mellanslag...

            booking_no = line[:8]
            all_details = line[:46]

            line = re.sub(r"\s{2}", "", line, 1)                # Letar upp första två mellanslagen räknat från vänster och tar bort dem
            
            if "F E E D E R   M A N I F E S T" in line:
                row_count += 8                                  # Om ovan finns i raden sätts row_count till 8 och tickar neråt. När det är på 0 körs nästa elif.

            elif row_count == 0:
                if line[19:30] == container:
                    hl_list[amount-1][5] = hl_list[amount-1][5] + float(line[89:97])    # Summerar nettovikten om containernummer är den samma som förra loopen.
                    hl_list[amount-1][8] = hl_list[amount-1][8] + int(line[74:78])      # Summerar packages ~

                else:
                    booking = int(line[:8])                                             # Om containernumret inte är densamma så sparas all data på nytt.
                    
                    if len(line[9:18].strip()) < 9:                                     # Om Work Order är under 9 symboler (när mellanslag är borta).
                        work_order = 0
                        container = line[17:21] + line[23:30]
                
                    else:                                                               
                        work_order = int(line[9:18])
                        container = line[19:30]

                    unit_type = line[31:35]
                    goods_desc = line[56:73].strip()
                    quantity = int(line[74:78])
                    net_weight = float(line[89:97])
                    vgm = float(line[98:106])
                    vessel = line[108:121].strip()
                    voyage = line[122:130].strip()
                    tod = tod_data
                    pod = pod_data
                
                    new_list = [booking, pod, tod, container, unit_type, net_weight, vgm,
                    work_order,quantity, goods_desc, vessel, voyage]

                    if hl_list != []:
                        if booking == hl_list[amount-1][0] and container == hl_list[amount-1][3]: # Om bokning och container är den samma ska nettovikterna summeras och
                            hl_list[amount-1][5] = hl_list[amount-1][5] + float(line[89:97])      # nästa loop påbörjas. När de inte längre är samma så fortsätter loopen
                            hl_list[amount-1][8] = hl_list[amount-1][8] + int(line[74:78])        # och listan new_list läggs till i hl_list.
                            new_list = []
                            continue
                    
                    
                    hl_list.append(new_list)
                    new_list = []
                    
                    amount += 1

            else:
                row_count -= 1
    return hl_list

def create_dataframe():

    df = pd.DataFrame(get_data())
    df.columns = ['BOOKING NUMBER', 'POL', 'TOL', 'CONTAINER', 'ISO TYPE', 'NET WEIGHT', 'VGM', 'PO NUMBER', 
    'PACKAGES', 'GOODS DESCRIPTION', 'OCEAN VESSEL', 'VOYAGE']

    df_info_sheet = pd.DataFrame(columns=['BOOKING NUMBER', 'MLO', 'POL', 'TOL', 'CONTAINER', 'ISO TYPE',
    'NET WEIGHT', 'POD STATUS', 'LOAD STATUS', 'VGM', 'OOG', 'REMARK', 'IMDG', 'UNNR', 'CHEM REF', 'MRN', 'TEMP',
    'PO NUMBER', 'CUSTOMS STATUS', 'PACKAGES', 'GOODS DESCRIPTION', 'OCEAN VESSEL', 'VOYAGE', 'ETA', 'FINAL POD'])

    df_info_sheet.loc[:,'BOOKING NUMBER'] = df['BOOKING NUMBER']
    df_info_sheet.loc[:, 'MLO'] = "HL"
    df_info_sheet.loc[:, 'POL'] = df['POL']
    df_info_sheet.loc[:, 'TOL'] = df['TOL']
    df_info_sheet.loc[:, 'CONTAINER'] = df['CONTAINER']
    df_info_sheet.loc[:, 'ISO TYPE'] = df['ISO TYPE']
    df_info_sheet.loc[:, 'NET WEIGHT'] = df['NET WEIGHT']
    df_info_sheet.loc[:, 'POD STATUS'] = "T"
    df_info_sheet.loc[:, 'LOAD STATUS'] = "LA"
    df_info_sheet.loc[:, 'VGM'] = df['VGM']
    df_info_sheet.loc[:, 'OOG'] = ""
    df_info_sheet.loc[:, 'REMARK'] = ""
    df_info_sheet.loc[:, 'IMDG'] = ""
    df_info_sheet.loc[:, 'UNNR'] = ""
    df_info_sheet.loc[:, 'CHEM REF'] = ""
    df_info_sheet.loc[:, 'MRN'] = ""
    df_info_sheet.loc[:, 'TEMP'] = ""
    df_info_sheet.loc[:, 'PO NUMBER'] = df['PO NUMBER']
    df_info_sheet.loc[:, 'CUSTOMS STATUS'] = "T1"
    df_info_sheet.loc[:, 'PACKAGES'] = df['PACKAGES']
    df_info_sheet.loc[:, 'GOODS DESCRIPTION'] = df['GOODS DESCRIPTION']
    df_info_sheet.loc[:, 'OCEAN VESSEL'] = df['OCEAN VESSEL']
    df_info_sheet.loc[:, 'VOYAGE'] = df['VOYAGE']
    df_info_sheet.loc[:, 'ETA'] = ""
    df_info_sheet.loc[:, 'FINAL POD'] = ""

    df_info_sheet = df_info_sheet.sort_values(by=['BOOKING NUMBER', 'CONTAINER'], ascending= (True, True)).copy()

    return df_info_sheet

def copy_data():

    dataframe = create_dataframe()

    wb = xw.Book.caller()
    ws = wb.sheets['RESULTAT']
    ws['A1'].options(pd.DataFrame, header=1, index=False).value = dataframe

def compare_FPOD_bokningsblad():

    bokningsblad = chosen_file()

    with xw.App(visible=False) as app:
        wb_bokningsblad = app.books.open(bokningsblad)
        ws_bokningsblad = wb_bokningsblad.sheets['INFO']
        range_bokningsblad = ws_bokningsblad.range('A4').expand() #dynamisk range
        df_bokningsblad = ws_bokningsblad.range(range_bokningsblad).options(pd.DataFrame, index=False, header=True).value #dynamisk range
        df_bokningsblad = df_bokningsblad[['BOOKING NUMBER', 'FINAL POD']].copy()
        wb_bokningsblad.close()

    wb_hl = xw.Book.caller()
    ws_hl = wb_hl.sheets['RESULTAT']

    lower_row = ws_hl.range('A' + str(ws_hl.cells.last_cell.row)).end('up').row

    df_hl = ws_hl.range('A1:Y' + str(lower_row)).options(pd.DataFrame, index=False, header=True).value #dynamisk range

    df_hl = df_hl['BOOKING NUMBER']
    df_hl = pd.DataFrame(df_hl).copy()

    dict_bokningsblad = dict(zip(df_bokningsblad['BOOKING NUMBER'], df_bokningsblad['FINAL POD']))
    df_hl['FINAL POD'] = df_hl['BOOKING NUMBER'].map(dict_bokningsblad)

    ws_hl.range('Y1').options(pd.Series, header=1, index=False).value = df_hl['FINAL POD']



if __name__ == '__main__':
    main()