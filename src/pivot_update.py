import pandas as pd
import xlwings as xw
import functions as fn

def main():
    
    def get_data():
        wb = xw.Book.caller()
        get_data.info_sheet = wb.sheets['INFO']
        get_data.data_sheet = wb.sheets['DATA']

        data_table = get_data.info_sheet.range('A4').expand()
        df = get_data.info_sheet.range(data_table).options(pd.DataFrame, index=False, header=True).value
        df = pd.DataFrame(df).copy()

        return df

    if get_data().shape[0] == 0:
        return
    else:
        update_data_sheet(get_data(), get_data.data_sheet)


def update_data_sheet(df: pd.DataFrame, data_sheet: xw.sheets):

    data_df = pd.DataFrame()

    df.loc[:, 'TARE'] = fn.get_tare(df)
    data_df.loc[:, 'get_max_weight'] = fn.get_max_weight(df)/1000
    data_df.loc[:, 'mlo'] = df['MLO']
    data_df.loc[:, 'tol'] = df['TOL']
    data_df.loc[:, "20feet"] = 0
    data_df.loc[:, "40feet"] = 0
    data_df.loc[:, 'get_teus'] = fn.get_TEUs(df)
    data_df.loc[data_df['get_teus'] == 1, "20feet"] = 1
    data_df.loc[data_df['get_teus'] == 2, "40feet"] = 1
    data_df.loc[:, 'mt_or_not'] = df['LOAD STATUS'].str.contains('MT')
    data_df.loc[data_df['mt_or_not'] == True, 'mt_teus'] = data_df['get_teus']
    data_df.loc[data_df['mt_or_not'] == False, 'la_teus'] = data_df['get_teus']
    data_df.loc[data_df['mt_or_not'] == True, 'mt_weight'] = df['TARE'] / 1000
    data_df.loc[data_df['mt_or_not'] == False, 'la_weight'] = data_df['get_max_weight']

    data_sheet.range('M4').options(pd.DataFrame, index=False, header=True).value = data_df.copy()