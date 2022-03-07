import pandas as pd
import xlwings as xw
import functions as fn

wb = xw.Book.caller()
info_sheet = wb.sheets['INFO']
data_sheet = wb.sheets['DATA']
data_table = info_sheet.range('A4').expand()
df = info_sheet.range(data_table).options(pd.DataFrame, index=False, header=True).value

def main():
    update_info_sheet()
    update_data_sheet()
    
def update_data_sheet():

    data_df = pd.DataFrame()

    df.loc[:, 'TARE'] = fn.get_tare(df)
    data_df.loc[:, 'MLO_check'] = fn.MLO_check(df)
    data_df.loc[:, 'terminal_check'] = fn.terminal_check(df)
    data_df.loc[:, 'cargo_type_check'] = fn.cargo_type_check(df)
    data_df.loc[:, 'container_check'] = df['CONTAINER'].apply(fn.container_check, 1)
    data_df.loc[:, 'load_status_check'] = fn.load_status_check(df)
    data_df.loc[:, 'reefer_check'] = fn.reefer_check(df)
    data_df.loc[:, 'dg_check'] = "TBA"
    data_df.loc[:, 'port_check'] = "TBA"
    data_df.loc[:, 'vessel_check'] = "TBA"
    data_df.loc[:, 'customs_check'] = "TBA"
    data_df.loc[:, 'get_max_weight'] = fn.get_max_weight(df)
    data_df.loc[:, 'get_TEUs'] = fn.get_TEUs(df)

    data_sheet.range('A5').options(pd.DataFrame, index=False, header=False).value = data_df

def update_info_sheet():

    terminal = ['tpl_terminal', 'TERMINAL OUTPUT', 'TOL']
    cargo_type = ['tpl_cargo_type', 'TYPE OUTPUT', 'ISO TYPE']
    vessel = ['tpl_vessels', 'HL VESSEL OUTPUT', 'OCEAN VESSEL']

    df.loc[:, 'TOL'] = fn.get_template_type(df, terminal)
    df.loc[:, 'ISO TYPE'] = fn.get_template_type(df, cargo_type)
    df.loc[:, 'OCEAN VESSEL'] = fn.get_template_type(df, vessel)
    df.loc[df['FINAL POD'] == "ZAZBA", 'FINAL POD'] = "ZADUR"

    info_sheet.range('D5').options(pd.Series, index=False, header=False).value = df['TOL']
    info_sheet.range('F5').options(pd.Series, index=False, header=False).value = df['ISO TYPE']
    info_sheet.range('Y5').options(pd.Series, index=False, header=False).value = df['FINAL POD']
