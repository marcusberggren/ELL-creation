import xlwings as xw
import pandas as pd
import numpy as np
import functions as fn
import create_ell as ce

def main():
    pre_coprar()

def pre_coprar():
    name = "PRE_COPRAR_"
    df = fn.get_caller_df().copy()
    df.loc[df['PACKAGES'].isnull(), 'PACKAGES'] = 1
    df_update = ce.work_with_df(df)
    df_dummy_added = creating_container_dummy(df_update)
    return fn.save_to_ell(name, ce.cargo_detail(df_dummy_added), ce.manifest(df_dummy_added))

def creating_container_dummy(df: pd.DataFrame):
    num_empty_rows = df['CONTAINER'].isnull().sum()
    num_as_text_in_list = np.arange(num_empty_rows).astype(str)
    list_filled_with_zeros = ["DUMY" + numbers.zfill(7) for numbers in num_as_text_in_list]
    df.loc[df['CONTAINER'].isnull(), 'CONTAINER'] = list_filled_with_zeros
    return df

if __name__ == '__main__':
    file_path = fn.get_mock_caller('0109_Bokningsblad.xlsb')
    xw.Book(file_path).set_mock_caller()
    pre_coprar()