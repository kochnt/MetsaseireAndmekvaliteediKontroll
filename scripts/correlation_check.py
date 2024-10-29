import numpy as np
def correlation_check(corr1, corr2, Dict):
    '''
    correlation_check() - Kontrolib, et esimene andmeelement oli kooskõlas teise elemendiga

    Parameetrid:       
    corr1: esimene andmeelement  
    corr2: teine andmeelement   
    Dict: sõnastik, milles asuvad elementide funktsionaalsed sõltuvused   

    Dimensioon: Reeglipärasus	   
    Probleemi liik: Funktsionaalse sõltuvuse rikkumine
    
    '''
    try:
        Key = corr1
        Value = corr2
        if isinstance(Key, (float, np.float32, np.float64)):  # Kontrollin, kas Key on arv
            if np.isnan(Key):  # Kontrollin, kas Key on NaN
                return True  
        if isinstance(Value, (float, np.float32, np.float64)):  # Kontrollin, kas Value on arv
            if np.isnan(Value):  # Kontrollin, kas Value on NaN
                return True  
        Key = str(Key)   
        Value = str(Value)
        # Kontrollin, kas võti on sõnastikus olemas ja kas väärtus vastab võtmele
        if Key in Dict and Value in Dict[Key]:
            return True
        else:
            return False

    except TypeError:
        return False
    
# Test
# test_dict = {'1000 okka kuivkaal': {np.nan},
#              'Aktiivne happesus': {'vaba H+'},
#              'Alumiinium': {'Al'},
#              'Ammooniumlämmastik (NH4N)': {'NH4N'},
#              'Asendusalumiinium': {'Al-ex'},
#              'Asendushappesus': {'Al+H'},
#              'Asenduskaalium': {'K-ex'},
#              'Asenduskaltsium': {'Ca-ex'},
#              'Asendusmagneesium': {'Mg-ex'},
#              'Asendusmangaan': {'Mn-ex'},
#              'Asendusnaatrium': {'Na-ex'},
#              'Asendusraud': {'Fe-ex'},
#              'Boor': {'B'},
#              'Boor (kõik puuliigid)': {'B'},
#              'Elavhõbe': {'Hg'},
#              'Elavhõbe (kõik puuliigid)': {'Hg'},
#              'Elektrijuhtivus': {np.nan},
#              'Fosfor': {'P'},
#              'Kaadmium': {'Cd'},
#              'Kaadmium (kõik puuliigid)': {'Cd'},
#              'Kaalium': {'K'},
#              'Kaalium (kõik puuliigid)': {'K'},
#              'Kahjustuse avaldumine': {np.nan},
#              'Kahjustuse koht võras': {np.nan},
#              'Kahjustuse põhjus': {np.nan},
#              'Kahjustuse põhjuse nimi': {np.nan},
#              'Kahjustuse sümptom': {np.nan},
#              'Kahjustuse ulatus': {np.nan},
#              'Kahjustuse vanus': {np.nan},
#              'Kaltsium': {'Ca'},
#              'Kaltsium (kõik puuliigid)': {'Ca'},
#              'Karbonaadid': {np.nan},
#              'Kloriid': {'Cl'},
#              'Kroom': {'Cr'},
#              'Kroom (kõik puuliigid)': {'Cr'},
#              'Käbikandvus': {np.nan},
#              'Ladva seisund': {np.nan},
#              'Lahustunud orgaaniline süsinik': {'DOC'},
#              'Leelisus': {np.nan, 'Alka'},
#              'Lisavõrsete hulk': {np.nan},
#              'Lõimis (tekstuurne klass WRB järgi)': {np.nan},
#              'Magneesium': {'Mg'},
#              'Magneesium (kõik puuliigid)': {'Mg'},
#              'Mangaan': {'Mn'},
#              'Mangaan (kõik puuliigid)': {'Mn'},
#              'Mullaniiskus': {np.nan},
#              'Mullavett proovis': {np.nan},
#              'Naatrium': {'Na'},
#              'Nikkel': {'Ni'},
#              'Nikkel (kõik puuliigid)': {'Ni'},
#              'Nitraatlämmastik (NO3N)': {'NO3N'},
#              'Okastiku vanuseklassid': {np.nan},
#              'Okka/lehekadu kogu võra ulatuses': {np.nan},
#              'Okka/lehekadu võra ülemises 1/3 osas': {np.nan},
#              'Orgaanilise kihi kuivkaal': {np.nan},
#              'Plii': {'Pb'},
#              'Plii (kõik puuliigid)': {'Pb'},
#              'Puistu I rinde keskmine vanuseklass': {np.nan},
#              'Puu kahjustatud osa': {np.nan},
#              'Puu kasvuklass': {np.nan},
#              'Puu kõrgus': {np.nan},
#              'Puu rinnasdiameeter': {np.nan},
#              'Puu seisund': {np.nan},
#              'Puu vanus': {np.nan},
#              'Puu vanuseklass': {np.nan},
#              'Raud': {'Fe'},
#              'Raud (kõik puuliigid)': {'Fe'},
#              'Sademete hulk': {np.nan},
#              'Sulfaatväävel (SO4S)': {'SO4S'},
#              'Tsink': {'Zn'},
#              'Tsink (kõik puuliigid)': {'Zn'},
#              'Vaadeldav võra osa': {np.nan},
#              'Vanimad okkad': {np.nan},
#              'Vanuse määramise meetod': {np.nan},
#              'Varise kuivkaal m² kohta': {np.nan},
#              'Varise kuivkaal m² kohta (kõik puuliigid)': {np.nan},
#              'Vask': {'Cu'},
#              'Vask (kõik puuliigid)': {'Cu'},
#              'Väävel': {'S'},
#              'Väävel (kõik puuliigid)': {'S'},
#              'Võra nähtavus': {np.nan},
#              'Võra varjutatus': {np.nan},
#              'Võrdluspuu': {np.nan},
#              'pH': {'pH'},
#              'pH (CaCl2)': {'pH (CaCl2)'},
#              'pH (H2O)': {'pH (H2O)'},
#              'Õitsemine': {np.nan},
#              'Üldfosfor': {'Püld'},
#              'Üldfosfor (kõik puuliigid)': {'Püld'},
#              'Üldlämmastik': {'Nüld'},
#              'Üldlämmastik (kõik puuliigid)': {'Nüld'},
#              'Üldorgaaniline süsinik': {'TOC'},
#              'Üldorgaaniline süsinik (kõik puuliigid)': {'TOC'}}

#print(correlation_check('Leelisus', np.nan, test_dict))


