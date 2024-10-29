library(readxl)

# Panen exceli leht datafreimi
row_metsaseire_2022 <- read_excel("~/Desktop/Keskkonnaagentuur/MetsaseireAndmekvaliteediKontroll/data/Metsaseire_2022_a.xlsx", sheet = "Metsaseire_2022")


colnames(row_metsaseire_2022) <- gsub(" ", "_", colnames(row_metsaseire_2022))

colnames(row_metsaseire_2022)