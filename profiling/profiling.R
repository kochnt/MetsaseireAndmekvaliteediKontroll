library(readxl) # Paket mis aitab lugeda excel
library(dlookr) # Paket profileerimise jaoks

# Panen exceli leht datafreimi
metsaseire_2022 <- read_excel("~/Desktop/Keskkonnaagentuur/MetsaseireAndmekvaliteediKontroll/data/Metsaseire_2022_a.xlsx", sheet = "Metsaseire_2022")

# Profileerimise aruanned
eda_web_report(metsaseire_2022)

diagnose_web_report(metsaseire_2022)

