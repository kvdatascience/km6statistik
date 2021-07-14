# Autor: Lars Kroll
# Datum: 2.7.2019

# angepasst by Beate Zoch-Lesniak
# Datum:02.12.2020
# neuer Pfad, neue Spalte KV, Daten für 2019 und 2020 ergänzt

library("readxl")
library("tidyverse")



km_6_path <- "G:/30 Fachbereich_3/14 Daten/KM6_Statistik"
#km_6_files <- list.files(km_6_path,pattern=".xls", full.names = T) # hatte damit irgendwie ein Problem mit Einlesen der 2018er Daten: Fehler: `path` does not exist: 'G:/30 Fachbereich_3/14 Daten/KM6_Statistik/~$KM6_2018.xlsx' 
km_6_files <- list.files(km_6_path,pattern="_20", full.names = T)

mynames <- read_xls(km_6_files[1] , col_names = F) %>%  # read first file
  filter(row_number() %in% c(3,5,6)) %>% # read labels
  gather(number,name) %>% group_by(number) %>%  # group by column
  mutate(lineno=paste0("line_",row_number())) %>% ungroup() %>% spread(lineno,name) %>% # change row labels to new columns
  mutate(i=as.numeric(str_remove_all(number,"[.]"))) %>%  arrange(i) %>%  # arrange by col
  fill(line_1,line_2) %>% # fill down
  mutate(name=str_replace_all(paste(line_1,line_2,line_3,sep=".")," ","_"),
         name=ifelse(i==1,"Region_und_Alter",name)) %>% 
  pull(name) # pull out all cleaned labels

mynames

km_6.Daten <- sapply(km_6_files, read_excel,  skip=7, simplify=FALSE,col_names = mynames) %>% 
  bind_rows(.id = "id") %>% 
  mutate(Jahr=str_remove(id,"G:/30 Fachbereich_3/14 Daten/KM6_Statistik/KM6_"),Jahr=as.numeric(str_split_fixed(Jahr,"[.]",n=2)[,1])) %>% 
  select(-id) %>% 
  mutate(Region=ifelse(is.na(Mitglieder.Insgesamt.Zusammen),Region_und_Alter,NA),
         Alter=ifelse(!is.na(Mitglieder.Insgesamt.Zusammen),Region_und_Alter,NA)) %>% 
  fill(Region) %>%  filter(!is.na(Alter)) %>% 
  select(-Region_und_Alter) %>% gather(Merkmal,Wert,contains(".")) %>% 
  mutate(Versichertengruppe=str_split_fixed(Merkmal,"[.]",n=3)[,1],
         Untergruppe=str_split_fixed(Merkmal,"[.]",n=3)[,2],
         Geschlecht=str_split_fixed(Merkmal,"[.]",n=3)[,3]) %>% 
  select(Jahr,Region,Versichertengruppe,Untergruppe,Geschlecht,Alter,"Anzahl"=Wert) %>%
  filter(Anzahl>0) %>% filter(Versichertengruppe!="Mitglieder_und_Familienangehörige_zusammen" & 
                                Untergruppe!="Insgesamt" &
                                Region != "Bund" &
                                Alter !="alle Altersgruppen" &
                                Geschlecht !="Zusammen")


write_excel_csv2(km_6.Daten,path = "G:/30 Fachbereich_3/14 Daten/KM6_Statistik/KM6_aufbereitet.csv")
saveRDS(km_6.Daten, file = "G:/30 Fachbereich_3/14 Daten/KM6_Statistik/KM6_aufbereitet.Rds")

regions_kv <- read_excel("G:/30 Fachbereich_3/14 Daten/KM6_Statistik/KV_Regionen_Zuordnung.xlsx")

km_6.Daten_plus_KV <- left_join(km_6.Daten, regions_kv, by = "Region")
km_6.Daten_plus_KV %>% filter(is.na("KV_Name")) #keine missings
km_6.Daten_plus_KV %>% filter(is.na("KV"))

write_excel_csv2(km_6.Daten_plus_KV,path = "G:/30 Fachbereich_3/14 Daten/KM6_Statistik/KM6_aufbereitet_plus_KV.csv")
saveRDS(km_6.Daten_plus_KV, file = "G:/30 Fachbereich_3/14 Daten/KM6_Statistik/KM6_aufbereitet_plus_KV.Rds")

