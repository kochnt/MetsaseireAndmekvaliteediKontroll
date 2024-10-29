
import pandas as pd
def column_check(value1, value2):
    '''
    Funktsioon kontrollib, kas ühes väärtuses on väärtus, aga teises mitte.
    Kui ühes on väärtus ja teises pole, siis tagastatakse True.
    Kui mõlemas on või pole väärtust, siis tagastatakse False.
    '''
    if (pd.notna(value1) and pd.isna(value2)) or (pd.isna(value1) and pd.notna(value2)):
        return True
    else:
        return False