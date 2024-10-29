import pandas as pd
import numpy as np

def unique_val_to_list(df, colName):
    """
    unique_val_to_list() - Eemaldab unikaalsed väärtused DataFrame veerust ja salvestab need nimekirja.

    Parameetrid:
    df: DataFrame, millest tuleb välja võtta unikaalsed väärtused.   
    colName: Veeru nimi, millest tuleb välja võtta unikaalsed väärtused.

    Tagastab unikaalsete väärtuste loend.
    """
    
    unique_val = df[colName].unique()

    print(len(unique_val))

    unique_val_list = unique_val.tolist()


    return unique_val_list

file_path = 'C:/Users/User/Desktop/MetsaseireAndmekvaliteediKontroll/data/Metsaseire_2022_a.xlsx'

df = pd.read_excel(file_path, sheet_name='Metsaseire_2022') # Dataframe kus asub 'Metsaseire_2022' tabel 
#unique_list = unique_val_to_list(df, 'element_kood')
#print(unique_list)

#column_list = ['Programm',	'Seiretöö_nimetus', 'Vastutav_partner', 'Vastutav_isik', 'Seirekoha_KKR', 
#                'Seirekoha_nimi', 'Seirekoha_staatus', 'Näitaja_nimetus', 'Vaatlusgrupp', 'Väärtuse_staatus', 'Väärtuse_täpsustus']

column_list = ['Mõõdetud_väärtuse_ühik', 'Väärtuse_ühik']
for i in column_list:
    unique_list = unique_val_to_list(df, i)
    print(unique_list)
    



