from datetime import datetime
import re
def format_seirekoha_kkr_check(seirekoha_kkr):
    '''
    format_seirekoha_kkr_check() - kontrollib, et andmeelement oli esitatud formaadis ABC1234567

    Parameetrid:   
    seirekoha_kkr - andmeelement

    Dimensioon: Reeglipärasus  
    Probleemi liik: Andmemustritest kõrvalekalded
    
    '''
    try:
        seirekoha_kkr_str = str(seirekoha_kkr)
        datetime_regex = re.compile(r'^[A-Z]{3}\d{7}$')

        result = bool(datetime_regex.match(seirekoha_kkr_str))

        return result
    except ValueError:
        return False

# Tests
#print(format_seirekoha_kkr_check('SJA8279000'))
    