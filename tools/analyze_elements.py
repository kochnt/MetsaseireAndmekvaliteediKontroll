import pandas as pd
import numpy as np

def analyze_elements(df):
    """
    Analüüsib DataFrame'i, et saada teavet maksimumide, miinimumide, mõõtühikute ja puuduva teabe kohta.
    
    Parameetrid:
    df: DataFrame, mis sisaldab veerge 'Näitaja_nimetus', 'Mõõdetud_väärtuse_ühik' ja 'Mõõdetud_arvväärtus'.
    
    Tagastab kaks sõnastikku: üks analüüsi tulemustega ja teine ilma väärtusteta elementidega.
    """
    # Rühmitame andmed 'Element' ja 'Ühikud' järgi
    grouped = df.groupby(['Näitaja_nimetus', 'Mõõdetud_väärtuse_ühik'])

    # Loome tühja sõnastiku tulemuste ja eraldi sõnastiku ilma väärtusteta elementide jaoks
    analysis_results = {}
    elements_without_values = {}

    # Itereerime rühmade üle
    for (element, unit), group in grouped:
        element_data = group['Mõõdetud_arvväärtus']
        
        # Filtreerime välja mitte-tühjad väärtused
        non_null_data = element_data.dropna()

        if not non_null_data.empty:
            max_value = non_null_data.max()
            min_value = non_null_data.min()
            if element not in analysis_results:
                analysis_results[element] = {}
            analysis_results[element][unit] = {
                'Maksimum': max_value,
                'Miinimum': min_value
            }
        else:
            if element not in elements_without_values:
                elements_without_values[element] = []
            elements_without_values[element].append(unit)

    return analysis_results, elements_without_values


file_path = '/Keskkonnaagentuur/MetsaseireAndmekvaliteediKontroll/data/Metsaseire_2022_a.xlsx'
df = pd.read_excel(file_path, sheet_name='Metsaseire_2022')

# Rakendame funktsiooni DataFrame'ile
results_with_values, results_without_values = analyze_elements(df)

# Tulemuste trükkimine elementide jaoks, millel on andmed
print("Elemendid andmetega:")
for element, units in results_with_values.items():
    print(f"{element}")
    for unit, values in units.items():
        print(f"  Mõõdetud väärtuse ühik: {unit}")
        print(f"    max: {values['Maksimum']}, min: 0")

# Ilma väärtusteta elementide trükkimine
print("Elemendid ilma väärtusteta:")
for element, units in results_without_values.items():
    print(f"Näitaja_nimetus: {element}")
    print(f"  Mõõtühikud ilma väärtusteta: {', '.join(units)}")
    print("-" * 30)
