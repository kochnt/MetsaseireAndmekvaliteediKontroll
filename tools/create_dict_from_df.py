import pandas as pd
from collections import defaultdict
from pprint import pprint

def create_dict_from_df(df, key_col, value_col):
    """
    Loob sõnastiku kahest DataFrame'i veerust, 
    kus iga võtme (key_col unikaalne väärtus)
    jaoks on vastav väärtuste loend (value_col unikaalsed väärtused)
    
    Argumendid:
    df: DataFrame
    key_col: veeru nimi, mida kasutatakse võtmetena
    value_col: veeru nimi, mida kasutatakse väärtustena
    
    Tagastab:
    Sõnastiku, kus võtmed on key_col unikaalsed väärtused 
    ja väärtused on loendid value_col unikaalsetest väärtustest.
    """
    result_dict = defaultdict(set)  # Kasutan set unikaalsete väärtuste hoidmiseks
    
    for _, row in df.iterrows():
        key = row[key_col]
        value = row[value_col]
        
        # Kontrollin, kas väärtuste kogumis on juba NaN ja lisan NaN ainult üks kord
        if pd.isna(value):
            if not any(pd.isna(v) for v in result_dict[key]):
                result_dict[key].add(value)
        else:
            result_dict[key].add(value)

    return result_dict
    
  
file_path = '/MetsaseireAndmekvaliteediKontroll/data/Metsaseire_2022_a.xlsx'
df = pd.read_excel(file_path, sheet_name='Metsaseire_2022')

result = create_dict_from_df(df, 'Mullaproovi_sügavus', 'Mullaproov')
pprint(result)
