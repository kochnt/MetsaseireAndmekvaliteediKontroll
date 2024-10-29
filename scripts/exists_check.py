import pandas as pd

def exists_check(value):
    '''
    exists_check() - kontrollib, kas kohustuslik atribut on väärtustatud.   

    Parameetrid:    
    value: andmeelement.

    Dimensioon: Täelikkus   
    Probleemi liik: Puuduv väärtus   

    '''
    if pd.isna(value):  # Kontrollib np.nan ja None
        return False
    if value in ["", [], {}]:  
        return False
    return True

# Test 

# print(exists_check(r))