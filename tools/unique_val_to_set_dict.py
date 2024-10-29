import pandas as pd

def unique_val_to_set_dict(df, column_list):
    """
    unique_val_to_set_dict() - Loob sõnastiku, kus võtmed on veergude nimed ja väärtused on unikaalsete väärtuste hulgad.

    Parameetrid:
    df: DataFrame, millest tuleb unikaalsed väärtused välja võtta.
    column_list: Veergude nimekiri, mille jaoks tuleb unikaalsete väärtuste sõnastik luua.

    Tagastab sõnastiku, kus võtmed on veergude nimed ja väärtused on unikaalsete väärtuste hulgad.
    """
    
    unique_dict = {}
    
    for col_name in column_list:
        unique_dict[col_name] = set(df[col_name].unique())
    
    return unique_dict

file_path = 'C:/Users/User/Desktop/MetsaseireAndmekvaliteediKontroll/data/Metsaseire_2022_a.xlsx'

df = pd.read_excel(file_path, sheet_name='Metsaseire_2022') 

column_list = ['Programm', 'Seiretöö_nimetus', 'Vastutav_partner', 'Vastutav_isik', 'Seirekoha_KKR', 
                'Seirekoha_nimi', 'Seirekoha_staatus', 'Näitaja_nimetus', 'Vaatlusgrupp', 'Väärtuse_staatus',
                 'Erimärk', 'Mõõdetud_väärtuse_ühik', 'Väärtuse_ühik', 'Analüüsimeetodi_standard',
       'Analüüsimeetodi_nimetus', 'Analüüsimeetodi_tööjuhendi_nr', 'Analüüsimeetodi_allikas', 'Väärtuse_täpsustus']

unique_values_dict = unique_val_to_set_dict(df, column_list)
print(unique_values_dict)

