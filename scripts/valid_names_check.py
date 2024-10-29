
def valid_names_check(vname, valid_names):
    '''
    valid_names_check() - kontrollib kõiki nimesid 
    ja kui element vastab etteantud loendile siis tagastab funktsioon True.

    Parameetrid:   
    vname: andmeelement.  
    validNames: õigete nimedega andmete massiiv(loend).  

    Tagastab:
    True või Falses

    Dimensioon: Õigsus
    Probleemi liik: Väärtusvahemiku rikkumine

    '''
    if vname in valid_names:
        return True # kui leiab nimi andme massiivis siis tagastab True
    else:
        return False
    
# Test
# validNames = ["Pakri", "Ristna"]

# print(valid_names_check('Pakri'))