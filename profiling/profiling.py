import os
import pandas as pd
from ydata_profiling import ProfileReport
import sweetviz as sv


# Exceli faili asukoht
file_path = "/Users/nathanaelkoch/Desktop/Keskkonnaagentuur/MetsaseireAndmekvaliteediKontroll/data/Metsaseire_2022_a.xlsx"
#"c:/AHPraktika/Andmeaidad/Kliima_Andmeait.xlsx"

#-------------------------------------------------------------------------------------------------------

# Teisendan Exceli lehed andmeraamideks 

forest_table = pd.read_excel(file_path, sheet_name='Metsaseire_2022')


#-------------------------------------------------------------------------------------------------------

# Loon aruanded iga andmeraami kohta

profile_forest_table = ProfileReport(forest_table, title="Profiling Report Metsaseire_2022")



#-------------------------------------------------------------------------------------------------------
# profile_list = [profile_element_table, profile_element_jaam_vaatlus_table, profile_element_kuu_table,
#                  profile_element_paev_table, profile_element_tund_table, profile_element_minut_table]
profile_list = [profile_forest_table]
# aruandefailide asukohad
base_path = '/Users/nathanaelkoch/Desktop/Keskkonnaagentuur/MetsaseireAndmekvaliteediKontroll/profiling'

output_list = [f'{base_path}/Metsaseire_2022_ProfileReport.html']
# output_list = [f'{base_path}/f_kliima_element_ProfileReport.html', f'{base_path}/f_kliima_jaam_vaatlus_ProfileReport.html',
#                 f'{base_path}/f_kliima_kuu_ProfileReport.html', f'{base_path}/f_kliima_paev_ProfileReport.html',
#                   f'{base_path}/f_kliima_tund_ProfileReport.html', f'{base_path}/f_kliima_minut_ProfileReport.html']
#-------------------------------------------------------------------------------------------------------

# kui selline aruandefail puudub, siis loon uus aruandefail
for i in range(len(output_list)):
    if not os.path.exists(output_list[i]):
        profile_list[i].to_file(output_file=output_list[i])
        print(f'{output_list[i]} on loodud')
    else:
        print(f'Fail on juba olemas: {output_list[i]}')

report = sv.analyze(forest_table)
report.show_html(f'{base_path}/Metsaseire_2022_sweetreport.html')
#print(elementTable.head())
