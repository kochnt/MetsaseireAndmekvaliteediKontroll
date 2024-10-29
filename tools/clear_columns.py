import pandas as pd

# Excelist andmete lugemine
df = pd.read_excel("/MetsaseireAndmekvaliteediKontroll/data/Metsaseire_2022_a.xlsx")

# Praeguste veergude nimede vaatamine
print("Enne:", df.columns)

# Tühjuste eemaldamine ja asendamine alljoonega
df.columns = df.columns.str.replace(' ', '_')

# Uute veergude nimede vaatamine
print("Pärast:", df.columns)

df.to_excel("/MetsaseireAndmekvaliteediKontroll/data/Metsaseire_2022_a.xlsx", index=False)
