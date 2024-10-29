
import pandas as pd
def condition_vaartus_check(arvvaartus, vaartus_muud):
    '''
    condition_vaartus_check() - funktioon kontrollib, et kui Arvväärtus väli on tühi, siis peab olema täidetud Väärtus (muud) väli
    
    Parameetrid:      
    arvvaartus: Arvväärtus väli 
    vaartus_muud: Väärtus (muu) väli
   
    Dimensioon: Õigsus  
    Probleemi liik: Väärtusvahemiku rikkumine

    '''
    if pd.isna(arvvaartus) and pd.notna(vaartus_muud):
        return True
    elif pd.isna(vaartus_muud) and pd.notna(arvvaartus):
        return True
    else:
        return False
    
