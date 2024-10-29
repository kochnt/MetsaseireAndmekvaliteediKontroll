def condition_seirekoha_y_check(seirekoha_y):
    '''
    condition_seirekoha_y_check() - funktioon kontrollib, et seirekoha ja seotud seirejaama y koordinaadid
    olid vahemikus 387377 ja 733508
    Parameetrid:      
    seirekoha_y: andmeelement 

    Dimensioon: Õigsus  
    Probleemi liik: Väärtusvahemiku rikkumine

    '''
    try:
        y = float(seirekoha_y)
        if 387377 <= seirekoha_y <= 733508:
            return True
        else:
            return False

    except (ValueError, TypeError):
        return False
# Test
#print(condition_seirekoha_y_check(733508))