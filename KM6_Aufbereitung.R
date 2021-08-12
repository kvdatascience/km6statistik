# Autor: Lars Kroll
# Datum: 2.7.2019

# angepasst by Beate Zoch-Lesniak
# Datum:02.12.2020
# neuer Pfad, neue Spalte KV, Daten für 2019 und 2020 erg?nzt

library("readxl")
library("tidyverse")

regions_kv <- read_excel("KV_Regionen_Zuordnung.xlsx")

km_6_files <- list.files(pattern="_20", full.names = T)

alldata <- tibble()

for (thefile in km_6_files ){
  mynames <- read_xls(km_6_files[1] , col_names = F) %>%  # read first file
    filter(row_number() %in% c(3,5,6)) %>% # read labels
    gather(number,name) %>% group_by(number) %>%  # group by column
    mutate(lineno=paste0("line_",row_number())) %>% ungroup() %>% spread(lineno,name) %>% # change row labels to new columns
    mutate(i=as.numeric(str_remove_all(number,"[.]"))) %>%  arrange(i) %>%  # arrange by col
    fill(line_1,line_2) %>% # fill down
    mutate(name=str_replace_all(paste(line_1,line_2,line_3,sep=".")," ","_"),
           name=ifelse(i==1,"Region_und_Alter",name)) %>% 
    pull(name) 
  alldata <- bind_rows(alldata,read_excel(thefile,  skip=7,col_names = mynames) %>% mutate(id=thefile))
}

km_6.Daten <- alldata %>%
  #Wert 0 in der Spalte Mitglieder.Insgesamt.Zusammen zu NA umkodieren,
  #da in einigen Jahren in Excel unsichtbare 0 in eigentlich leeren Feldern enthalten ist
  #da ansonsten bei den Mitgliedern insgesamt zusammen in keiner Altersgruppe 0 Personen
  #sein sollten, erscheint mir das als der einfachste Weg der Korrektur
  mutate(Mitglieder.Insgesamt.Zusammen = na_if(Mitglieder.Insgesamt.Zusammen, 0)) %>% 
  mutate(Region=ifelse(is.na(Mitglieder.Insgesamt.Zusammen),Region_und_Alter,NA),
         Alter=ifelse(!is.na(Mitglieder.Insgesamt.Zusammen),Region_und_Alter,NA)) %>% 
  fill(Region) %>%  filter(!is.na(Alter)) %>% 
  select(-Region_und_Alter) %>% gather(Merkmal,Wert,contains(".")) %>% 
  mutate(Versichertengruppe=str_split_fixed(Merkmal,"[.]",n=3)[,1],
         Untergruppe=str_split_fixed(Merkmal,"[.]",n=3)[,2],
         Geschlecht=str_split_fixed(Merkmal,"[.]",n=3)[,3]) %>% 
  select(id,Region,Versichertengruppe,Untergruppe,Geschlecht,Alter,"Anzahl"=Wert) %>%
  filter(Anzahl>0) %>% filter(Versichertengruppe!="Mitglieder_und_Familienangehörige_zusammen" & 
                                Untergruppe!="Insgesamt" &
                                Region != "Bund" &
                                Alter !="alle Altersgruppen" &
                                Geschlecht !="Zusammen") %>%
  mutate(Jahr=as.numeric(str_split_fixed(str_remove(id,"./KM6_"),"\\.",2)[,1])) %>%
  left_join(.,regions_kv,by="Region") 

km_6.Daten.agg <-km_6.Daten %>% group_by(Jahr,KV_Name,KV,Versichertengruppe,Untergruppe,Geschlecht,Alter) %>%
  summarise(Anzahl =sum(Anzahl,na.rm=FALSE)) %>% 
  ungroup()


write_excel_csv2(km_6.Daten,path ="KM6_aufbereitet.csv")
saveRDS(km_6.Daten, file = "KM6_aufbereitet.Rds")

#check
km_6_agg_KV_Jahr <- km_6.Daten %>% 
  group_by(KV, KV_Name, Jahr) %>% 
  summarise(Anzahl = sum(Anzahl)) %>% 
  arrange(KV_Name,Jahr)

ggplot(km_6_agg_KV_Jahr, aes(x=Jahr, y=Anzahl/100000, color=as.factor(KV_Name), group=as.factor(KV_Name))) +
  geom_line() +
  scale_y_continuous(limits = c(0,150))

km_6_agg_Jahr <- km_6.Daten %>% 
  group_by(Jahr) %>% 
  summarise(Anzahl = sum(Anzahl))
