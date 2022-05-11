import pandas as pd
import xlwings as xw
import os
import functions as fn

def main():
    pre_export()

def pre_export():
    template_file = fn.get_path('tpl_pre_export_continent')
    df = fn.get_caller_df().copy()
    df_group = df.groupby('MLO')

    vessel = fn.get_caller_df.vessel
    voyage = fn.get_caller_df.voyage
    pol = fn.get_caller_df.pol

    wb_caller_path = xw.Book.caller().fullname
    folder_path_bokningsblad = os.path.split(wb_caller_path)[0]
    file_name = "PRE_EXPORT_" + vessel + "_" + str(voyage[:5]) + "_" + pol + "_"   #utelämna slutet för att kompletteras i loop nedan
    name_of_file_and_path = os.path.join(folder_path_bokningsblad, file_name)


    for name, group in df_group:
        with xw.App(visible=False) as app:
            wb = app.books.open(template_file)
            wb.save(name_of_file_and_path + name + ".xlsx")
            sheet = wb.sheets['INFO']
            sheet.range('A2').value = vessel
            sheet.range('B2').value = voyage
            sheet.range('C2').value = pol
            sheet.range('A5').options(pd.DataFrame, index=False, header=False).value = group
            wb.save()
            wb.close()

if __name__ == '__main__':
    file_path = fn.get_mock_caller('0109_Bokningsblad.xlsb')
    xw.Book(file_path).set_mock_caller()
    pre_export()