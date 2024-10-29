def condition_seirekoha_x_check(seirekoha_x):
    '''
    condition_seirekoha_x_check() - funktioon kontrollib, et seirekoha ja seotud seirejaama x koordinaadid
    olid vahemikus 6383851 ja 6606656
    Parameetrid:      
    seirekoha_x: andmeelement 
   
    Dimensioon: Õigsus  
    Probleemi liik: Väärtusvahemiku rikkumine

    '''
    try:
        x = float(seirekoha_x)
        if 6383851 <= seirekoha_x <= 6606656:
            return True
        else:
            return False

    except (ValueError, TypeError):
        return False
    
# Test
#print(condition_seirekoha_x_check(6606656))