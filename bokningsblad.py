import pandas as pd
import xlwings as xw
import functions as fn

wb = xw.Book.caller()
info_sheet = wb.sheets['INFO']
data_sheet = wb.sheets['DATA']
data_table = info_sheet.range('A4').expand()
df = info_sheet.range(data_table).options(pd.DataFrame, index=False, header=True).value

csv_main_path = r'\BOLLORE\XPF - Documents\MAINTENANCE\templates\stored-data-'

def main():
    update_info_sheet()
    update_data_sheet()
    
def update_data_sheet():

    data_df = pd.DataFrame()

    df['TARE'] = fn.get_tare(df)
    data_df['MLO_check'] = fn.MLO_check(csv_main_path, df)
    data_df['terminal_check'] = fn.terminal_check(csv_main_path, df)
    data_df['cargo_type_check'] = fn.cargo_type_check(csv_main_path, df)
    data_df['container_check'] = df['CONTAINER'].apply(fn.container_check, 1)
    data_df['load_status_check'] = fn.load_status_check(csv_main_path, df)
    data_df['reefer_check'] = fn.reefer_check(df)
    data_df['dg_check'] = "TBA"
    data_df['port_check'] = "TBA"
    data_df['vessel_check'] = "TBA"
    data_df['customs_check'] = "TBA"
    data_df['get_max_weight'] = fn.get_max_weight(df)
    data_df['get_TEUs'] = fn.get_TEUs(df)

    data_sheet.range('A5').options(pd.DataFrame, index=False, header=False).value = data_df

def update_info_sheet():

    terminal = ['template-terminal', 'TERMINAL OUTPUT', 'TOL']
    cargo_type = ['template-cargo-type', 'TYPE OUTPUT', 'ISO TYPE']
    vessel = ['template-vessels', 'HL VESSEL OUTPUT', 'OCEAN VESSEL']

    df.loc[:, 'TOL'] = fn.get_template_type(csv_main_path, df, terminal)
    df.loc[:, 'ISO TYPE'] = fn.get_template_type(csv_main_path, df, cargo_type)
    df.loc[:, 'OCEAN VESSEL'] = fn.get_template_type(csv_main_path, df, vessel)

    info_sheet.range('D5').options(pd.Series, index=False, header=False).value = df['TOL']
    info_sheet.range('F5').options(pd.Series, index=False, header=False).value = df['ISO TYPE']
    info_sheet.range('V5').options(pd.Series, index=False, header=False).value = df['OCEAN VESSEL']