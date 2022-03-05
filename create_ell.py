import xlwings as xw
import pandas as pd
from functions import get_csv_data, regex_no_extra_whitespace, get_tare, get_max_weight
import numpy as np
from datetime import datetime
import os
from pathlib import Path



def main():
    create_ell()

def define_path():

    csv_main_path = r'\BOLLORE\XPF - Documents\MAINTENANCE\templates\stored-data-'


    return csv_main_path

def copy_sheets_to_workbook(df1: pd.DataFrame, df2: pd.DataFrame, vessel, voyage, leg, pol):

    wb_caller_path = xw.Book.caller().fullname
    folder_path_bokningsblad = os.path.split(wb_caller_path)[0]
    time_str = datetime.now().strftime("%y%m%d")
    ell_file_name = "ELL_" + vessel + "_" + str(voyage[:5]) + "_" + pol + "_" + time_str + ".xlsx"
    name_of_file_and_path = os.path.join(folder_path_bokningsblad, ell_file_name)

    home = Path.home()
    ell_template_path = str(home) + r'\BOLLORE\XPF - Documents\MAINTENANCE\templates\template-ell.xlsx'
    
    with xw.App(visible=False) as app:
        wb = app.books.open(ell_template_path)
        cargo_detail_sheet = wb.sheets['Cargo Detail']
        manifest_sheet = wb.sheets['Manifest']
        cargo_detail_sheet.range('A6').options(pd.DataFrame, index=False, header=False).value = df1.copy()
        cargo_detail_sheet.range('A2').value = vessel
        cargo_detail_sheet.range('B2').value = voyage
        cargo_detail_sheet.range('C2').value = leg
        cargo_detail_sheet.range('F2').value = pol
        manifest_sheet.range('A2').options(pd.DataFrame, index=False, header=False).value = df2.copy()
        wb.save(name_of_file_and_path)
        wb.close()

def create_ell():

    wb = xw.Book.caller()
    sheet = wb.sheets('INFO')
    data_table = sheet.range('A4').expand()
    df = sheet.range(data_table).options(pd.DataFrame, index=False, header=True).value

    vessel = sheet.range('A2').value
    voyage = str(sheet.range('B2').value)
    leg = sheet.range('C2').value
    pol = sheet.range('D2').value
    
    #Skapar 4 olika data frames från CSV-filer
    csv_main_path = define_path()

    df_cargo_type = get_csv_data(csv_main_path, 'cargo-type').copy()
    df_country = get_csv_data(csv_main_path, 'country').copy()
    df_mlo = get_csv_data(csv_main_path, 'mlo').copy()
    df_ocean_vessel = get_csv_data(csv_main_path, 'ocean-vessel').copy()
    

    # Regex som byter ut white spaces i början och slutet på varje instans i dataframe
    df = regex_no_extra_whitespace(df).copy()
    df_cargo_type = regex_no_extra_whitespace(df_cargo_type).copy()
    df_country = regex_no_extra_whitespace(df_country).copy()
    df_mlo = regex_no_extra_whitespace(df_mlo).copy()

    def work_with_df(df: pd.DataFrame):
        
        # Sätter ihop ISO TYPE och LOAD STATUS
        df.loc[:, 'ISO TYPE'] = df['ISO TYPE'].astype(str)
        df.loc[:, 'ISO STATUS'] = df['ISO TYPE'] + df['LOAD STATUS']

        # När NET WEIGHT är mindre än 100 men större än 0 multiplicera med 1000
        df.loc[(df['NET WEIGHT'] < 100) & (df['NET WEIGHT'] != 0), 'NET WEIGHT'] *= 1000
        
        # Lägger till "CHEM" och "MT" om boolean sann
        df.loc[df['IMDG'].notnull(), 'CHEM'] = "CHEM"
        mt_check = df['LOAD STATUS'].str.contains("MT")
        df.loc[mt_check, 'LOAD STATUS'] = "MT"

        # Ändrar allt som inte är "MT" till "LA"
        df.loc[df['LOAD STATUS'] != "MT", 'LOAD STATUS'] = "LA"

        # Ändrar alla instanser av ZAZBA till ZADUR
        change_pod = df['FINAL POD'] == "ZAZBA"
        df.loc[change_pod, 'FINAL POD'] = "ZADUR"

        # Ändrar MLO till MSK om HSL skeppar tomma enheter
        df.loc[(df['MLO'] == "HSL") & (df['LOAD STATUS'].str.contains("MT")), 'MLO'] = "MSK"

        # Bokningsnummer blir PO-nummer om det inte finns PO-nummer angett
        df.loc[:, 'PO NUMBER'] = np.where(df['PO NUMBER'].isnull(), df['BOOKING NUMBER'], df['PO NUMBER'])

        # Skapar ny kolumn med gods + MRN. Om MRN är noll (isnull) så läggs enbart gods till
        df.loc[:, 'GOODS+MRN'] = np.where(df['MRN'].isnull(), df['GOODS DESCRIPTION'], df['GOODS DESCRIPTION'] + " " + df['MRN'])

        # Lägger till 'T6 CARGO TYPE' till df när ISO STATUS matchar
        df = df.merge(df_cargo_type, on='ISO STATUS', how='left').copy()

        # Lägger till 'CALL SIGN' till df när OCEAN VESSEL matchar
        df = df.merge(df_ocean_vessel, how='left', on='OCEAN VESSEL').copy()

        # Ändrar 'PORT' till 2 bokstäver, används i merge #1 nedan
        df.loc[:, 'PORT'] = df['POL'].str[:2]

        #Merge df_country och 'PORT'
        df = df.merge(df_country, on='PORT', how='left').copy()
        df = df.merge(df_mlo, on='MLO', how='left').copy()

        # Viktigt att denna ändring görs för merge #2
        df.loc[:, 'PORT'] = pol[:2]

        # Merge #2 lägger till suffixes när kolumnerna 'COUNTRY' krockar
        df = df.merge(df_country, on='PORT', how='left', suffixes=('_CONSIGNEE','_SHIPPER')).copy()

        # Slår ihop flera conditions och sätter taravikterna rätt i df['TARE']
        df.loc[:, 'TARE'] = get_tare(df)

        """
        Beskrivning av get_max_weight:
        - summerar nettovikt och tara om VGM är tom och NET WEIGHT > 100
        - om NET WEIGHT < 100 och inte 0 multipliceras värdet med 1000
        - om VGM > 0 skrivs maxvärdet ut av NET WEIGHT och VGM
        """
        df.loc[:, 'MAX WEIGHT'] = get_max_weight(df) / 1000   # div med 1000 för tonvikt

        # Lägger till vikt i 'VGM-LA' om inte "MT" i kolumn
        mt_check = df['LOAD STATUS'] != "MT"
        df.loc[mt_check, 'VGM-LA'] = df['MAX WEIGHT']

        # Ytterligare två conditions men som skapar df['TRANSHIPMENT']
        conditions_pod =[
            df['POD STATUS'] == "T",
            df['POD STATUS'] == "Y"
            ]
        values_pod = ["Y", "N"]

        df['TRANSHIPMENT'] = np.select(conditions_pod, values_pod)
        return df
    
    df = work_with_df(df).copy()

    def cargo_detail(df: pd.DataFrame):
        df_cargo_detail = pd.DataFrame(columns=['Pod', 'Pod call seq', 'Pod terminal', 'Pod Status', 'POL',
        'Pol terminal', 'Pol Status', 'Shunted Terminal', 'Slot Owner', 'Slot Account', 'MLO', 'MLO PO',
        'Booking Reference', 'Ex Vessel', 'Ex Voyage', 'Next Vessel', 'Next Voyage', 'Mother Vessel',
        'Mother Vessel CallSign', 'Mother Voyage', 'POT', 'F.Destination', 'VIA', 'VIA terminal', 'Cargo type',
        'ISO Container Type', 'User Container Type', 'Commodity', 'OOG', 'Container No', 'Weight in MT', 'Stowage',
        'Door Open', 'Slot Killed', 'V.Type', 'Fr.Group', 'TempMax', 'TempMin', 'TempOpt', 'IMCO', 'FP', 'UN',
        'PSA Class', 'IMO Name', 'Chem Name', 'Remarks', 'OOH(CM)', 'OLF(CM)', 'OLA(CM)', 'OWP(CM)', 'OWS(CM)',
        'VGM Weight in MT', 'VGM Cert Signatory', 'VGM Certificate No', 'VGM Weighing Method',
        'VGM Cert Issuing Party', 'VGM Cert Issuing Address', 'VGM Cert Issue Date'])

        df_cargo_detail['Pod'] = df['POL']
        df_cargo_detail['Pod call seq'] = 1
        df_cargo_detail['Pod terminal'] = df['TOL']
        df_cargo_detail['Pod Status'] = "T"
        df_cargo_detail['POL'] = pol
        df_cargo_detail['Pol terminal'] = pol
        df_cargo_detail['Pol Status'] = "L"
        df_cargo_detail['Slot Owner'] = "XCL"
        df_cargo_detail['Slot Account'] = "XCL"
        df_cargo_detail['MLO'] = df['MLO']
        df_cargo_detail['MLO PO'] = df['PO NUMBER']
        df_cargo_detail['Booking Reference'] = df['BOOKING NUMBER']
        df_cargo_detail['Mother Vessel'] = df['OCEAN VESSEL']
        df_cargo_detail['Mother Vessel CallSign'] = df['CALL SIGN']
        df_cargo_detail['Mother Voyage'] = df['VOYAGE']
        df_cargo_detail['F.Destination'] = df['FINAL POD']
        df_cargo_detail['Cargo type'] = df['T6 CARGO TYPE']
        df_cargo_detail['ISO Container Type'] = df['ISO TYPE']
        df_cargo_detail['User Container Type'] = df['ISO TYPE']
        df_cargo_detail['Commodity'] = df['LOAD STATUS']
        df_cargo_detail['Container No'] = df['CONTAINER']
        df_cargo_detail['Weight in MT'] = df['MAX WEIGHT']
        df_cargo_detail['IMCO'] = df['IMDG']
        df_cargo_detail['UN'] = df['UNNR']
        df_cargo_detail['IMO Name'] = df['CHEM'] ## ADD?
        df_cargo_detail['Remarks'] = df['CHEM REF']
        df_cargo_detail['VGM Weight in MT'] = df['VGM-LA']
        return df_cargo_detail

    def manifest(df: pd.DataFrame):

        df_manifest = pd.DataFrame(columns=['Pod Terminal', 'MLO', 'B/L No', 'MLO PO', 'Booking Reference',
        'OBL Reference', 'Marks & Nos', 'No of Cntr', 'Type', 'Stc', 'No of Packages', 'Unit', 'Goods Desc',
        'Cargo Status', 'Transhipment', 'Seal No', 'Tare Weight in Kilos', 'Net Weight in Kilos',
        'Deep Sea Vessel', 'ETA', 'Rcvr', 'Shipper', 'Consignee', 'Notify', 'Product ID', 'Volume (Meter)',
        'Marks', 'MRN Number', 'Remarks', 'Port of Origin', 'Export PO', 'Package Content'])

        df_manifest['Pod Terminal'] = df['POL']
        df_manifest['MLO'] = df['MLO']
        df_manifest['MLO PO'] = df['PO NUMBER']
        df_manifest['Booking Reference'] = df['BOOKING NUMBER']
        df_manifest['Marks & Nos'] = df['CONTAINER']
        df_manifest['No of Cntr'] = 1
        df_manifest['Type'] = df['T6 CARGO TYPE']
        df_manifest['Stc'] = "STC"
        df_manifest['No of Packages'] = df['PACKAGES']
        df_manifest['Unit'] = "PK"
        df_manifest['Goods Desc'] = df['GOODS+MRN']
        df_manifest['Cargo Status'] = df['CUSTOMS STATUS']
        df_manifest['Transhipment'] = df['TRANSHIPMENT']
        df_manifest['Tare Weight in Kilos'] = df['TARE']
        df_manifest['Net Weight in Kilos'] = df['NET WEIGHT']
        df_manifest['Deep Sea Vessel'] = df['OCEAN VESSEL']
        df_manifest['ETA'] = df['ETA']
        df_manifest['Shipper'] = df['SHIPPER'] + " " + df['COUNTRY_SHIPPER']
        df_manifest['Consignee'] = df['CONSIGNEE'] + " " + df['COUNTRY_CONSIGNEE']
        df_manifest['Notify'] = df_manifest['Consignee']
        df_manifest['MRN Number'] = df['MRN']
        df_manifest['Package Content'] = df['GOODS DESCRIPTION']
        return df_manifest

    return copy_sheets_to_workbook(cargo_detail(df), manifest(df), vessel, voyage, leg, pol)
    

if __name__ == '__main__':
    xw.Book(r'C:\Users\SWV224\BOLLORE\XPF - Documents\0109_Bokningsblad_TEST_2-BSEGOTL116844.xlsb')

    create_ell()