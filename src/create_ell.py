import xlwings as xw
import pandas as pd
import functions as fn
import numpy as np


def main():
    create_ell()  

def work_with_df(df: pd.DataFrame):

    # Regex som byter ut white spaces i början och slutet på varje instans i dataframe
    df = fn.regex_no_extra_whitespace(df).copy()

    df.loc[:, 'ISO TYPE'] = df['ISO TYPE'].astype(str)
    df.loc[:, 'ISO STATUS'] = df['ISO TYPE'] + df['LOAD STATUS']
    df.loc[:, 'GOODS DESCRIPTION'] = df['GOODS DESCRIPTION'].astype(str)
    
    # Lägger till "CHEM" och "MT" om boolean sann
    df.loc[df['IMDG'].notnull(), 'CHEM'] = "CHEM"
    mt_check = df['LOAD STATUS'].str.contains("MT")
    df.loc[mt_check, 'LOAD STATUS'] = "MT"

    # Ändrar allt som inte är "MT" till "LA"
    df.loc[df['LOAD STATUS'] != "MT", 'LOAD STATUS'] = "LA"

    # Ändrar alla instanser av ZAZBA till ZADUR
    change_pod = df['FINAL POD'] == "ZAZBA"
    df.loc[change_pod, 'FINAL POD'] = "ZAPLZ"

    # Ändrar MLO till MSK om HSL skeppar tomma enheter
    df.loc[(df['MLO'] == "HSL") & (df['LOAD STATUS'].str.contains("MT")), 'MLO'] = "MSK"

    # Bokningsnummer blir PO-nummer om det inte finns PO-nummer angett
    df.loc[:, 'PO NUMBER'] = np.where(df['PO NUMBER'].isnull(), df['BOOKING NUMBER'], df['PO NUMBER'])

    # Skapar ny kolumn med gods + MRN. Om MRN är noll (isnull) så läggs enbart gods till
    df.loc[:, 'GOODS+MRN'] = np.where(df['MRN'].isnull(), df['GOODS DESCRIPTION'], df['GOODS DESCRIPTION'] + " " + df['MRN'])

    df.loc[:, 'T6 CARGO TYPE'] = fn.get_template_type(df, ['cargo_type', 'T6 CARGO TYPE', 'ISO STATUS'])

    df.loc[:, 'CALL SIGN'] = fn.get_template_type(df, ['ocean_vessel', 'CALL SIGN', 'OCEAN VESSEL' ])

    df.loc[:, 'POD SHORT'] = df['POL'].str[:2]
    df.loc[:, 'POL SHORT'] = fn.get_caller_df.pol[:2]
    df.loc[:, 'POD MANIFEST'] = fn.get_template_type(df, ['country', 'COUNTRY', 'POD SHORT'])
    df.loc[:, 'POL MANIFEST'] = fn.get_template_type(df, ['country', 'COUNTRY', 'POL SHORT'])
    df.loc[:, 'MLO MANIFEST'] = fn.get_template_type(df, ['mlo', 'NAME', 'MLO'])
    df.loc[:, 'TARE'] = fn.get_tare(df)

    """
    Beskrivning av get_max_weight:
    - summerar nettovikt och tara om VGM är tom och NET WEIGHT > 100
    - om NET WEIGHT < 100 och inte 0 multipliceras värdet med 1000
    - om VGM > 0 skrivs maxvärdet ut av NET WEIGHT och VGM
    """
    df.loc[:, 'MAX WEIGHT'] = fn.get_max_weight(df) / 1000   # div med 1000 för tonvikt

    # Lägger till vikt i 'VGM-LA' om inte "MT" i kolumn
    mt_check_boolean = df['LOAD STATUS'] != "MT"
    df.loc[mt_check_boolean, 'VGM-LA'] = df['MAX WEIGHT']

    # Ytterligare två conditions men som skapar df['TRANSHIPMENT']
    conditions_pod =[
        df['POD STATUS'] == "T",
        df['POD STATUS'] == "L"
        ]
    values_pod = ["Y", "N"]

    df['TRANSHIPMENT'] = np.select(conditions_pod, values_pod)
    return df


def cargo_detail(df: pd.DataFrame):
    df_cd = pd.DataFrame(columns=['Pod', 'Pod call seq', 'Pod terminal', 'Pod Status', 'POL',
    'Pol terminal', 'Pol Status', 'Shunted Terminal', 'Slot Owner', 'Slot Account', 'MLO', 'MLO PO',
    'Booking Reference', 'Ex Vessel', 'Ex Voyage', 'Next Vessel', 'Next Voyage', 'Mother Vessel',
    'Mother Vessel CallSign', 'Mother Voyage', 'POT', 'F.Destination', 'VIA', 'VIA terminal', 'Cargo type',
    'ISO Container Type', 'User Container Type', 'Commodity', 'OOG', 'Container No', 'Weight in MT', 'Stowage',
    'Door Open', 'Slot Killed', 'V.Type', 'Fr.Group', 'TempMax', 'TempMin', 'TempOpt', 'IMCO', 'FP', 'UN',
    'PSA Class', 'IMO Name', 'Chem Name', 'Remarks', 'OOH(CM)', 'OLF(CM)', 'OLA(CM)', 'OWP(CM)', 'OWS(CM)',
    'VGM Weight in MT', 'VGM Cert Signatory', 'VGM Certificate No', 'VGM Weighing Method',
    'VGM Cert Issuing Party', 'VGM Cert Issuing Address', 'VGM Cert Issue Date'])

    df_cd.loc[:, 'Pod'] = df['POL']
    df_cd.loc[:, 'Pod call seq'] = 1
    df_cd.loc[:, 'Pod terminal'] = df['TOL']
    df_cd.loc[:, 'Pod Status'] = df['POD STATUS']
    df_cd.loc[:, 'POL'] = fn.get_caller_df.pol
    df_cd.loc[:, 'Pol terminal'] = fn.get_template_type(df_cd, ['terminal', 'TERMINAL', 'POL'])
    df_cd.loc[:, 'Pol Status'] = "L"
    df_cd.loc[:, 'Slot Owner'] = "XCL"
    df_cd.loc[:, 'Slot Account'] = "XCL"
    df_cd.loc[:, 'MLO'] = df['MLO']
    df_cd.loc[:, 'MLO PO'] = df['PO NUMBER']
    df_cd.loc[:, 'Booking Reference'] = df['BOOKING NUMBER']
    df_cd.loc[:, 'Mother Vessel'] = df['OCEAN VESSEL']
    df_cd.loc[:, 'Mother Vessel CallSign'] = df['CALL SIGN']
    df_cd.loc[:, 'Mother Voyage'] = df['VOYAGE']
    df_cd.loc[:, 'F.Destination'] = df['FINAL POD']
    df_cd.loc[:, 'Cargo type'] = df['T6 CARGO TYPE']
    df_cd.loc[:, 'ISO Container Type'] = df['ISO TYPE']
    df_cd.loc[:, 'User Container Type'] = df['ISO TYPE']
    df_cd.loc[:, 'Commodity'] = df['LOAD STATUS']
    df_cd.loc[:, 'Container No'] = df['CONTAINER']
    df_cd.loc[:, 'Weight in MT'] = df['MAX WEIGHT']
    df_cd.loc[:, 'TempMax'] = df['TEMP']
    df_cd.loc[:, 'TempMin'] = df['TEMP']
    df_cd.loc[:, 'TempOpt'] = df['TEMP']
    df_cd.loc[:, 'IMCO'] = df['IMDG']
    df_cd.loc[:, 'UN'] = df['UNNR']
    df_cd.loc[:, 'IMO Name'] = df['CHEM'] ## ADD?
    df_cd.loc[:, 'Remarks'] = df['CHEM REF']
    df_cd.loc[:, 'VGM Weight in MT'] = df['VGM-LA']
    return df_cd

def manifest(df: pd.DataFrame):

    df_man = pd.DataFrame(columns=['Pod Terminal', 'MLO', 'B/L No', 'MLO PO', 'Booking Reference',
    'OBL Reference', 'Marks & Nos', 'No of Cntr', 'Type', 'Stc', 'No of Packages', 'Unit', 'Goods Desc',
    'Cargo Status', 'Transhipment', 'Seal No', 'Tare Weight in Kilos', 'Net Weight in Kilos',
    'Deep Sea Vessel', 'ETA', 'Rcvr', 'Shipper', 'Consignee', 'Notify', 'Product ID', 'Volume (Meter)',
    'Marks', 'MRN Number', 'Remarks', 'Port of Origin', 'Export PO', 'Package Content'])

    df_man.loc[:, 'Pod Terminal'] = df['POL']
    df_man.loc[:, 'MLO'] = df['MLO']
    df_man.loc[:, 'MLO PO'] = df['PO NUMBER']
    df_man.loc[:, 'Booking Reference'] = df['BOOKING NUMBER']
    df_man.loc[:, 'Marks & Nos'] = df['CONTAINER']
    df_man.loc[:, 'No of Cntr'] = 1
    df_man.loc[:, 'Type'] = df['T6 CARGO TYPE']
    df_man.loc[:, 'Stc'] = "STC"
    df_man.loc[:, 'No of Packages'] = df['PACKAGES']
    df_man.loc[:, 'Unit'] = "PK"
    df_man.loc[:, 'Goods Desc'] = df['GOODS+MRN']
    df_man.loc[:, 'Cargo Status'] = df['CUSTOMS STATUS']
    df_man.loc[:, 'Transhipment'] = df['TRANSHIPMENT']
    df_man.loc[:, 'Tare Weight in Kilos'] = df['TARE']
    df_man.loc[:, 'Net Weight in Kilos'] = fn.get_net_weight(df)
    df_man.loc[:, 'Deep Sea Vessel'] = df['OCEAN VESSEL']
    df_man.loc[:, 'ETA'] = df['ETA']
    df_man.loc[:, 'Shipper'] = df['MLO MANIFEST'] + " " + df['POL MANIFEST']
    df_man.loc[:, 'Consignee'] = df['MLO MANIFEST'] + " " + df['POD MANIFEST']
    df_man.loc[:, 'Notify'] = df_man['Consignee']
    df_man.loc[:, 'MRN Number'] = df['MRN']
    df_man.loc[:, 'Package Content'] = df['GOODS DESCRIPTION']
    return df_man

def create_ell():
    name = "ELL_"
    df = fn.get_caller_df().copy()
    df = work_with_df(df).copy()

    return fn.save_to_ell(name, cargo_detail(df), manifest(df))


if __name__ == '__main__':
    file_path = fn.get_mock_caller('0109_Bokningsblad.xlsb')
    xw.Book(file_path).set_mock_caller()
    create_ell()