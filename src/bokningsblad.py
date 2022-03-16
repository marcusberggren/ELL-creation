import pandas as pd
import xlwings as xw
import functions as fn



def main():
    wb = xw.Book.caller()
    info_sheet = wb.sheets['INFO']
    data_sheet = wb.sheets['DATA']

    data_table = info_sheet.range('A4').expand()
    df = info_sheet.range(data_table).options(pd.DataFrame, index=False, header=True).value
    df = pd.DataFrame(df).copy()
    
    if df.shape[0] == 0:
        return
    else:
        update_info_sheet(df, info_sheet)
        update_data_sheet(df, data_sheet)
    
def update_data_sheet(df: pd.DataFrame, data_sheet: xw.sheets):
    df = fn.regex_no_extra_whitespace(df)
    data_df = pd.DataFrame()

    df.loc[:, 'TARE'] = fn.get_tare(df)
    data_df.loc[:, 'MLO_check'] = fn.MLO_check(df)
    data_df.loc[:, 'terminal_check'] = fn.terminal_check(df)
    data_df.loc[:, 'container_check'] = df['CONTAINER'].apply(fn.container_check, 1)
    data_df.loc[:, 'cargo_type_check'] = fn.cargo_type_check(df)
    data_df.loc[:, 'load_status_check'] = fn.load_status_check(df)
    data_df.loc[:, 'oog_check'] = fn.oog_check(df)
    data_df.loc[:, 'dg_check'] = fn.dg_check(df)
    data_df.loc[:, 'reefer_check'] = fn.reefer_check(df)
    data_df.loc[:, 'po_number_check'] = fn.po_number_check(df)
    data_df.loc[:, 'customs_check'] = fn.customs_check(df)
    data_df.loc[:, 'vessel_check'] = fn.vessel_check(df)
    data_df.loc[:, 'fpod_check'] = fn.fpod_check(df)
    data_df.loc[:, 'get_max_weight'] = fn.get_max_weight(df)
    data_df.loc[:, 'get_TEUs'] = fn.get_TEUs(df)

    data_sheet.range('A4').options(pd.DataFrame, index=False, header=True).value = data_df

def update_info_sheet(df: pd.DataFrame, info_sheet: xw.sheets):

    df = fn.regex_no_extra_whitespace(df)

    mlo = ['ever_partner_code', 'EVER MLO', 'MLO']
    terminal = ['tpl_terminal', 'TERMINAL OUTPUT', 'TOL']
    cargo_type = ['tpl_cargo_type', 'TYPE OUTPUT', 'ISO TYPE']
    vessel = ['tpl_vessels', 'HL VESSEL OUTPUT', 'OCEAN VESSEL']
    fpod = ['tpl_ports', 'UNLOCODE', 'FINAL POD']

    df.loc[:, 'TOL'] = fn.get_template_type(df, terminal)
    df.loc[:, 'ISO TYPE'] = fn.get_template_type(df, cargo_type)
    df.loc[:, 'OCEAN VESSEL'] = fn.get_template_type(df, vessel)
    df.loc[:, 'FINAL POD'] = fn.get_template_type(df, fpod)
    df.loc[:, 'MLO'] = fn.get_template_type(df, mlo)

    info_sheet.range('A5').options(pd.DataFrame, index=False, header=False).value = df.copy()