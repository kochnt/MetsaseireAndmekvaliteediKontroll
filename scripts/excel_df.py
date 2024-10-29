import pandas as pd

def excel_df(table):
    '''
    Teisendan Excel leht andmeraamideks 
    '''
    file_path = '/Users/nathanaelkoch/Desktop/Keskkonnaagentuur/MetsaseireAndmekvaliteediKontroll/data/Metsaseire_2022_a.xlsx'

    try:
        df = pd.read_excel(file_path, sheet_name=table)
        return df
    except ValueError:
        return False
    
#print(excel_df('Metsaseire_2022'))

