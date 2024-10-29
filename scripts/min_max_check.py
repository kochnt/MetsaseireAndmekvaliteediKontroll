import pandas as pd

def min_max_check(nimetus, arvvaartus, uhik, limits):
    """
    min_max_check() - funktsioon kontrollib, kas mõõdetud väärtus jääb antud elemendi ja mõõtühiku lubatud vahemikku.
    
    Parameetrid:
    nimetus: Elemendi nimi
    arvvaartus: Mõõdetud arvväärtus
    uhik: Mõõtühik 
    limits: Sõnastik
    
    Dimensioon: Õigsus  
    Probleemi liik: Väärtusvahemiku rikkumine
    """
    
    # Kui elemendi nime pole, tagastatakse False
    if pd.isna(nimetus):
        return False
    
    if pd.isna(arvvaartus) and pd.isna(uhik):
        return True
    max_val = limits.get(nimetus, {}).get(uhik, {}).get('max')

    # Kui elemendi nimi on olemas, otsitakse seda limiitide sõnastikust
    if nimetus in limits:

        if uhik in limits[nimetus]:
            min_val = limits[nimetus][uhik]['min']
            max_val = limits[nimetus][uhik]['max']
            # Kontrollin, kas mõõdetud väärtus jääb lubatud vahemikku
            if min_val <= arvvaartus <= max_val:
                return True
            else:
                return False
        # Kui elemendi jaoks on määratud vahemik ilma mõõtühikuta (nt pH)
        elif None in limits[nimetus]:
            min_val = limits[nimetus][None]['min']
            max_val = limits[nimetus][None]['max']
            if min_val <= arvvaartus <= max_val:
                return True
            else:
                return False
        else:
            return False
    else:
        return False
