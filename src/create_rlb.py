import xlwings as xw
import pandas as pd
import os
from datetime import datetime
import functions as fn

def main():
    collecting_data()

def collecting_data():
    df = fn.get_caller_df()
    df.dropna(subset=['TEMP'], inplace=True)

    df_rlb = pd.DataFrame(columns=[
        'Reference', 'MLO', 'Terminal', 'Container No', 'Size',
        'Weight','Commodity', 'Temp Set', 'POD'])

    df_rlb.loc[:, 'Reference'] = df['BOOKING NUMBER']
    df_rlb.loc[:, 'MLO'] = df['MLO']
    df_rlb.loc[:, 'Terminal'] = df['TOL']
    df_rlb.loc[:, 'Container No'] = df['CONTAINER']
    df_rlb.loc[:, 'Size'] = df['ISO TYPE']
    df_rlb.loc[:, 'Weight'] = df['VGM']
    df_rlb.loc[:, 'Commodity'] = df['GOODS DESCRIPTION']
    df_rlb.loc[:, 'Temp Set'] = df['TEMP']
    df_rlb.loc[:, 'POD'] = df['POL']
    return finish(df_rlb)

def finish(df: pd.DataFrame):

    vessel = fn.get_caller_df.vessel
    voyage = fn.get_caller_df.voyage
    pol = fn.get_caller_df.pol
    len_df = len(df) - 1

    wb_caller_path = xw.Book.caller().fullname
    folder_path_bokningsblad = os.path.split(wb_caller_path)[0]
    time_str = datetime.now().strftime("%y%m%d")
    rlb_file_name = vessel + "_" + str(voyage) + "_" + pol + "_REEFER_LOG_BOOK_" + time_str + ".xlsx"
    rlb_file_name_pdf = vessel + "_" + str(voyage) + "_" + pol + "_REEFER_LOG_BOOK_" + time_str
    name_of_file_and_path = os.path.join(folder_path_bokningsblad, rlb_file_name)
    pdf_name = os.path.join(folder_path_bokningsblad, rlb_file_name_pdf)

    with xw.App(visible=False) as app:
            wb = app.books.open(fn.get_path('tpl_rlb'))
            wb.save(name_of_file_and_path)

            dcm_sheet = wb.sheets['RLB']
            dcm_sheet.range('C8').value = vessel
            dcm_sheet.range('G8').value = voyage
            dcm_sheet.range('I8').value = pol
            dcm_sheet.range((12, 1), (11 + len_df, 10)).insert('down')
            dcm_sheet.range((11, 1), (11, 10)).delete('up')
            dcm_sheet.range('B11').options(pd.DataFrame, index=False, header=False).value = df.copy()

            wb.save()
            wb.to_pdf(pdf_name)
            wb.close()


if __name__ == '__main__':
    file_path = fn.get_mock_caller('0109_Bokningsblad.xlsb')
    xw.Book(file_path).set_mock_caller()
    collecting_data()