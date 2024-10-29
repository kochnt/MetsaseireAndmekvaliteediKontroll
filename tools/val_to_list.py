import pandas as pd
def val_to_list(df, colName):
    """
    val_to_list() - Eemaldab kõik väärtused DataFrame veerust ja salvestab need nimekirja.

    Parameetrid:
    df: DataFrame, millest tuleb välja võtta kõik väärtused.   
    colName: Veeru nimi, millest tuleb välja võtta kõik väärtused.

    Tagastab väärtuste loend.
    """
    
    print(len(df[colName]))

    val_list = df[colName].tolist()

    return val_list


file_path = '/KliimaAndmekvaliteediKontroll/data/Kliima_Andmeait.xlsx'

df = pd.read_excel(file_path, sheet_name='f_kliima_element') # Dataframe kus asub 'jaam_kood' tabel jaam_vaatlus

val_list = val_to_list(df, 'element_yhik_eng')
print(val_list)


