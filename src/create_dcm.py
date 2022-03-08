import xlwings as xw
import pandas as pd
from datetime import datetime, date
import os
import functions as fn

def main():
    collecting_data()

def collecting_data():
    df = fn.get_caller_df().copy()

    #gather EMS-information and merges with df
    df_ems = fn.get_csv_data('ems').copy()
    df_ems = df_ems.rename(columns={'UNNO':'UNNR'})
    df = df.merge(df_ems, how='left', on='UNNR').copy()

    df = fn.regex_no_extra_whitespace(df).copy()
    df_ems = fn.regex_no_extra_whitespace(df_ems).copy()

    df.dropna(subset=['IMDG'], inplace=True)

    df_dg = pd.DataFrame(columns=[
        'MLO', 'Reference', 'TOL', 'SIZE', 'CONTAINER#', 'IMO class', 'UN no', 'IMO Name',
        'Package Group', 'MP', 'FP (°C)', 'NO. OF PK', 'Packages ', 'Gross weight ( kg )',
        'Net weight ( kg )', 'EMS', 'POD', 'Acceptance ref'
        ])

    df_dg.loc[:, 'MLO'] = df['MLO']
    df_dg.loc[:, 'Reference'] = df['BOOKING NUMBER']
    df_dg.loc[:, 'TOL'] = df['TOL']
    df_dg.loc[:, 'SIZE'] = df['ISO TYPE']
    df_dg.loc[:, 'CONTAINER#'] = df['CONTAINER']
    df_dg.loc[:, 'IMO class'] = df['IMDG']
    df_dg.loc[:, 'UN no'] = df['UNNR']
    df_dg.loc[:, 'IMO Name'] = ""
    df_dg.loc[:, 'Package Group'] = ""
    df_dg.loc[:, 'MP'] = ""
    df_dg.loc[:, 'FP (°C)'] = ""
    df_dg.loc[:, 'NO. OF PK'] = ""
    df_dg.loc[:, 'Packages'] = ""
    df_dg.loc[:, 'Gross weight ( kg )'] = ""
    df_dg.loc[:, 'Net weight ( kg )'] = ""
    df_dg.loc[:, 'EMS'] = df['EMS']
    df_dg.loc[:, 'POD'] = df['FINAL POD']
    df_dg.loc[:, 'Acceptance ref'] = df['CHEM REF']
   
    return finish(df_dg)

def finish(df: pd.DataFrame):

    vessel = fn.get_caller_df.vessel
    voyage = fn.get_caller_df.voyage
    pol = fn.get_caller_df.pol
    today = date.today().strftime("%Y-%m-%d")
    len_df = len(df)-1

    wb_caller_path = xw.Book.caller().fullname
    folder_path_bokningsblad = os.path.split(wb_caller_path)[0]
    time_str = datetime.now().strftime("%y%m%d")
    dgm_file_name = "DG_" + vessel + "_" + str(voyage[:5]) + "_" + pol + "_" + time_str + ".xlsx"
    name_of_file_and_path = os.path.join(folder_path_bokningsblad, dgm_file_name)

    with xw.App(visible=False) as app:
        wb = app.books.open(fn.get_path('tpl_dcm'))
        wb.save(name_of_file_and_path)

        dcm_sheet = wb.sheets['DCM']
        dcm_sheet.range('C11').value = vessel
        dcm_sheet.range('F11').value = voyage
        dcm_sheet.range('I11').value = pol
        dcm_sheet.range('F18').value = today
        dcm_sheet.range((14, 1), (13 + len_df, 19)).insert('down')
        dcm_sheet.range('B14').options(pd.DataFrame, index=False, header=False).value = df.copy()

        wb.save(name_of_file_and_path)
        wb.close()

if __name__ == '__main__':
    file_path = fn.get_mock_caller('0109_Bokningsblad.xlsb')
    xw.Book(file_path).set_mock_caller()
    collecting_data()
    