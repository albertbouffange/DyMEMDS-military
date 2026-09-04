library(tidyverse)
# library(viridis)
# library(ggpubr)
# library(hrbrthemes)
library(zoo)
library(readxl)
library(openxlsx)


#setwd("/home/albert/Nextcloud/R/IEA")
#ou
setwd("/home/abouffan/Nextcloud/R/rDMC")

### ============================================================== ###
###           Lecture et  nettoyage de la base rDMC                ###
### ============================================================== ###


# rdmc_france_raw<-read_csv("rDMC_raw_v1.csv", locale = locale(encoding = "UTF-8"))%>% 
#   filter(.[[3]] == "france"&service!="paramilitary") %>% 
#   mutate(equipment_type=str_replace(equipment_type,
#                                     "anti-tank/anti-infrastructure",
#                                     "anti-tank and anti-infrastructure"),
#          equipment_subtype=str_replace_all(equipment_subtype,
#                                            c("helicopter/aircraft carriers" = "helicopter and aircraft carriers",
#                                              "anti-aircraft/command cruiser" = "anti-aircraft and command cruiser")
#          ))
# 
# rdmc_france=rdmc_france_raw %>% 
#   group_by(equipment_type,unit_name) %>% #suppression de l'interpolation entre années manquantes
#   complete(year = full_seq(year, 1),
#            fill = list(unit_count = 0)) %>%
#   group_by(year,equipment_type,unit_name) %>% #regroupement des différentes valeurs d'équipment_type pour une même année (éviter la dent de scie)
#   summarise(unit_count = sum(unit_count))
# rdmc_france |> write_csv("rDMC_France.csv")
rdmc_france=read_csv("rDMC_France.csv")



### nettoyage des données ### ----

charger_correspondance <- function(fichier) {
  read_excel("rdmc_correspondances.xlsx", sheet = fichier, col_names = FALSE)
  read.csv(fichier, header = FALSE, stringsAsFactors = FALSE) %>%
    rename(canonique = V1) %>%
    pivot_longer(cols = -canonique, 
                 names_to = NULL, 
                 values_to = "unit_name", 
                 values_drop_na = TRUE) %>%
    filter(unit_name != "") %>%
    select(unit_name, canonique) %>%
    distinct()
}

appliquer_correspondance <- function(df, mapping) {
  df %>%
    left_join(mapping, by = "unit_name") %>%
    mutate(unit_name = coalesce(canonique, unit_name)) %>%
    select(-canonique)
}

regrouper <- function(df) {
  df %>%
    group_by(across(-unit_count)) %>%
    summarise(unit_count = sum(unit_count, na.rm = TRUE), .groups = "drop")
}


exporter_dymemds <- function(df, nom_fichier_dymemds) {
  df %>% #export pour dymemds
    # group_by(year) %>%
    # summarise(unit_count=sum(unit_count)) %>%
    pivot_wider(names_from = year, values_from = unit_count) %>%
    write_excel_csv(paste0(nom_fichier_dymemds, "_dymemds.xlsx"))
}


### hélicos

helico_final <- rdmc_france %>%
  filter(equipment_type == "helicopters") %>%
  appliquer_correspondance(charger_correspondance("helis correspondance")) %>%
  filter(unit_name %notin% c("H-34", "Lynx", "Bell")) %>% #dernier filtrage
  regrouper() %>%
  mutate(unit_count=if_else(unit_name=="Gazelle"&year%in%c("1987","1990"),NA,unit_count), #corrections
         unit_count=if_else(unit_name=="Puma"&year%in%c("1976"),NA,unit_count),
         unit_count=if_else(unit_name=="Alouette II"&year%in%c("1987","1988","1989"),NA,unit_count),
         unit_count=if_else(unit_name=="Alouette"&year%in%c("1976","1977","1983","1984","1987","1988"),NA,unit_count)) #|> 
  # helico_final |> exporter_dymemds("helicopters")


# helico_final %>%  ggplot(aes(x=year,y=unit_count,fill=unit_name))+
#   geom_area() +
#   theme_bw()+
#   labs(title="helicopters")+
#   theme(legend.position="bottom")
# ggsave("clean_helicopters.png")
# 
# write_csv(helico_final, "clean_helicopters.csv")



###tanks

tanks_final <- rdmc_france %>%
  filter(equipment_type == "armoured fighting vehicles") %>%
  appliquer_correspondance(charger_correspondance("tanks correspondance")) %>%
  filter(unit_name!="NA",
         unit_name%notin%c("M-47", "AMX", "VP-90","Aravis","Buffalo","AMX-1-ECH","BvS-10","Leclerc DNG"), #séries vides/inutiles
  ) |> 
  regrouper() %>%
  mutate(unit_count=if_else(unit_name=="VAB"&year%in%c("1983","1990","2005","2006","2007","2008"),NA,unit_count),
         unit_count=if_else(unit_name=="AMX-30"&year%in%c("1997","2005","2006","2007","2008","2011"),NA,unit_count),
         unit_count=if_else(unit_name=="AMX-13"&year%in%c("1981","1983"),NA,unit_count),
         unit_count=if_else(unit_name=="AMX VTT"&year%in%c("1982","1990","1991"),NA,unit_count),
         unit_count=if_else(unit_name=="VBL M-11"&year%in%c("2003","2004"),NA,unit_count)
         ) #|> 
  # exporter_dymemds()

# tanks_final %>%  ggplot(aes(x=year,y=unit_count,fill=unit_name))+
#   geom_area() +
#   theme_bw()+
#   theme(legend.position="bottom")+
#   labs(title="armoured fighting vehicles")
# ggsave("clean_armoured fighting vehicles.png")
# write_csv(tanks_final, "clean_armoured fighting vehicles.csv")


### avions
avions_final <- rdmc_france %>%
  filter(equipment_type == "aircraft") %>%
  appliquer_correspondance(charger_correspondance("avions correspondance")) %>%
  filter(unit_name!="héli",
         unit_name%notin%c("A-1D Skyraider", "DC-4", "Etendard", "Mystere IV", "Br 765","15C-47D","T-33/Mystere IV/Noratlas") #dernier filtrage dont séries vides
  ) |> 
  regrouper() %>%
  mutate(unit_count=if_else(unit_name=="Transall C-160"&year%in%c("1996"),NA,unit_count),
         unit_count=if_else(year%in%c("1988","1990","1992","1994"),NA,unit_count),
         unit_count=if_else(unit_name=="Mirage III"&year%in%c("1977","1986"),NA,unit_count),
         unit_count=if_else(unit_name=="light training"&year%in%c("1985"),NA,unit_count)
         ) #|> 
# exporter_dymemds()

# avions_final %>%  ggplot(aes(x=year,y=unit_count,fill=unit_name))+
#   geom_area() +
#   theme_bw()+
#   theme(legend.position="bottom")+
#   labs(title="aircraft")+
#   theme(legend.position="bottom",
#         legend.text = element_text(size=6),
#         legend.key.size = unit(0.2, 'cm'))+
#   guides(fill=guide_legend(ncol=4,byrow=TRUE))
# ggsave("clean_aircraft.png")
# write_csv(avions_final, "clean_aircraft.csv")


### principal surface
principal_final <- rdmc_france %>%
  filter(equipment_type == "principal surface combatants") %>%
  appliquer_correspondance(charger_correspondance("principal surface correspondance")) %>%
  filter(unit_name!="NA",
         unit_name!="1.0",
         unit_name%notin%c("Forbin","Aquitaine") #dernier filtrage
  ) |> 
  regrouper() %>%
  mutate(unit_count=if_else(unit_name=="Type F-70"&year%in%c("2000"),NA,unit_count),
         unit_count=if_else(unit_name=="F-67"&year%in%c("1995","2000"),NA,unit_count),
         unit_count=if_else(unit_name=="D'Estienne d'Orves"&year%in%c("1994"),NA,unit_count),
         unit_count=if_else(unit_name=="Floreal"&year%in%c("2000"),NA,unit_count),
         unit_count=if_else(unit_name=="Escorteurs"&year%in%c("1982"),NA,unit_count),
         unit_count=if_else(unit_name=="Suffren"&year%in%c("2000"),NA,unit_count)
         ) #|> 
# exporter_dymemds()

# principal_final %>%  ggplot(aes(x=year,y=unit_count,fill=unit_name))+
#   geom_area() +
#   theme_bw()+
#   theme(legend.position="bottom")+
#   labs(title="principal surface combatants")+
#   theme(legend.position="bottom",
#         legend.text = element_text(size=6),
#         legend.key.size = unit(0.2, 'cm'))+
#   guides(fill=guide_legend(ncol=4,byrow=TRUE))
# ggsave("clean_principal surface.png")
# write_csv(principal_final, "clean_principal surface.csv")

### patrol and coastal 
patrol_final <- rdmc_france %>%
  filter(equipment_type == "patrol and coastal combatants") %>%
  appliquer_correspondance(charger_correspondance("patrol and coastal correspondance")) %>%
  filter(unit_name!='NA',
         unit_name%notin%c("Stellis","Mercure","La Combattante","Le Fougueux","Mercure","Iris","Phenix","Laperouse","Gowind") #dernier filtrage
  ) |> 
  regrouper() %>%
  mutate(unit_count=if_else(unit_name=="P-400"&year%in%c("1994"),NA,unit_count),
         unit_count=if_else(unit_name=="Trident"&year%in%c("1986","1987","1988","1989","1990","1991","1992"),4,unit_count)) %>% 
  add_row(year=1983,equipment_type="patrol and coastal combatants", unit_name="Leopard (Instruction)",unit_count=8) %>% 
  add_row(year=1984,equipment_type="patrol and coastal combatants", unit_name="Leopard (Instruction)",unit_count=8) %>% 
  add_row(year=1985,equipment_type="patrol and coastal combatants", unit_name="Leopard (Instruction)",unit_count=8) %>% 
  add_row(year=1986,equipment_type="patrol and coastal combatants", unit_name="Leopard (Instruction)",unit_count=8) %>% 
  add_row(year=1987,equipment_type="patrol and coastal combatants", unit_name="Leopard (Instruction)",unit_count=8) #|> 
# exporter_dymemds()

# patrol_final %>%  ggplot(aes(x=year,y=unit_count,fill=unit_name))+
#   geom_area() +
#   theme_bw()+
#   theme(legend.position="bottom")+
#   labs(title="patrol and coastal")+
#   theme(legend.position="bottom",
#         legend.text = element_text(size=8),
#         legend.key.size = unit(0.3, 'cm'))
# ggsave("clean_patrol and coastal.png")
# write_csv(patrol_final, "clean_patrol and coastal.csv")

### submarines
sub_final <- rdmc_france %>%
  filter(equipment_type == "submarines") %>%
  appliquer_correspondance(charger_correspondance("submarines correspondance")) %>%
  filter(unit_name!="Exocet",
         unit_name!="M-4/TN-70",
         unit_name!="M-20/TN-60"
         ) |> 
  regrouper() %>%
  mutate(unit_count=if_else(unit_name=="Le Redoutable"&year%in%c("1991"),6,unit_count),
         unit_count=if_else(unit_name=="Le Redoutable"&year%in%c("1992"),5,unit_count),
         unit_count=if_else(unit_name=="Daphne"&year%in%c("1990"),NA,unit_count),
         unit_count=if_else(unit_name=="Rubis"&year%in%c("1991"),5,unit_count),
         unit_count=if_else(unit_name=="Le Triomphant"&year%in%c("1996"),0,unit_count),
         unit_count=if_else(unit_name=="Le Triomphant"&year%in%c("2004","2005","2006","2007","2008","2009","2010"),3,unit_count),
         unit_count=if_else(unit_name=="Le Triomphant"&year%in%c("2011","2012","2013"),4,unit_count)) %>% 
  # add_row(year=1971, unit_name="Le Redoutable",unit_count=1) %>% 
  # add_row(year=1972, unit_name="Le Redoutable",unit_count=1) %>% 
  # add_row(year=1973, unit_name="Le Redoutable",unit_count=2) %>% 
  # add_row(year=1974, unit_name="Le Redoutable",unit_count=3) %>% 
  # add_row(year=1975, unit_name="Le Redoutable",unit_count=3) %>% 
  # add_row(year=1976, unit_name="Le Redoutable",unit_count=4) %>% 
  # add_row(year=1977, unit_name="Le Redoutable",unit_count=4) %>% 
  # add_row(year=1978, unit_name="Le Redoutable",unit_count=4) %>% 
  add_row(year=1979, equipment_type="submarines", unit_name="Le Redoutable",unit_count=4) %>% 
  add_row(year=1980, equipment_type="submarines", unit_name="Le Redoutable",unit_count=5) %>% 
  add_row(year=1981, equipment_type="submarines", unit_name="Le Redoutable",unit_count=5) %>% 
  add_row(year=1982, equipment_type="submarines", unit_name="Le Redoutable",unit_count=5) %>% 
  add_row(year=1983, equipment_type="submarines", unit_name="Le Redoutable",unit_count=5) %>% 
  add_row(year=1984, equipment_type="submarines", unit_name="Le Redoutable",unit_count=5) %>% 
  add_row(year=1985, equipment_type="submarines", unit_name="Le Redoutable",unit_count=6) %>% 
  add_row(year=1986, equipment_type="submarines", unit_name="Le Redoutable",unit_count=6) %>% 
  add_row(year=1987, equipment_type="submarines", unit_name="Le Redoutable",unit_count=6) #|> 
# exporter_dymemds()


# sub_final %>%  ggplot(aes(x=year,y=unit_count,fill=unit_name))+
#   geom_area() +
#   theme_bw()+
#   theme(legend.position="bottom")+
#   labs(title="submarines")
# ggsave("clean_submarines.png")
# write_csv(sub_final, "clean_submarines.csv")

sub_final %>% #export pour dymemds
  # group_by(year) %>% 
  # summarise(unit_count=sum(unit_count)) %>% 
  pivot_wider(names_from=year,values_from = unit_count) %>%
  write_excel_csv("submarines_dymemds.xlsx")



### ============================================================== ###
###  Export DyMEMDS : construction de StocksMilitaryMods.xlsx      ###
### ============================================================== ###



### 1. Paramètres du format cible ### ----

ANNEES_OBS  <- 1970:2014   # colonnes de données
ANCRAGE_BAS <- 1900        # colonne d'ancrage à 0 avant mise en service
ANCRAGE_HAUT<- 2100        # colonne d'ancrage à 0 après 2014 (Vensim interpole)

ANNEES_COL <- c(ANCRAGE_BAS, ANNEES_OBS, ANCRAGE_HAUT)

# feuille -> préfixe du code Vensim (= élément du subscript MilitaryTypeMod)
PREFIXES <- c(
  Armoured    = "MilitaryArmoured",
  Aircrafts   = "MilitaryAircrafts",
  Helicopters = "MilitaryHelicopters",
  Patrol      = "MilitaryPatrol",
  Surface     = "MilitarySurface",
  Submarines  = "MilitarySubmarines"
)


### 2. Fonctions utilitaires ### ----

# "SP-2H Neptune" -> "Sp2HNeptune" ; "A400M Atlas" -> "A400MAtlas" ;
# "D'Estienne d'Orves" -> "DEstienneDOrves" ; "AlphaJet" -> "Alphajet"
# Règle : découpage sur tout séparateur non alphanumérique ; dans chaque jeton,
# une lettre est mise en minuscule si et seulement si elle suit une autre lettre.
nom_vensim <- function(x) {
  jeton_camel <- function(jeton) {
    ch <- strsplit(jeton, "")[[1]]
    if (!length(ch)) return("")
    est_lettre <- grepl("[A-Za-z]", ch)
    prec_lettre <- c(FALSE, head(est_lettre, -1))
    ifelse(est_lettre & prec_lettre, tolower(ch), toupper(ch)) %>% paste(collapse = "")
  }
  x %>%
    iconv(to = "ASCII//TRANSLIT") %>%          # accents -> ASCII
    str_split("[^A-Za-z0-9]+") %>%
    map_chr(~ .x %>% discard(~ .x == "") %>% map_chr(jeton_camel) %>% paste(collapse = ""))
}

# Étape commune aux deux versions : complète la grille (matériel x année),
# retrouve les bornes de service de chaque matériel.
completer_grille <- function(df, annees_obs) {
  long <- df %>%
    ungroup() %>%
    select(unit_name, year, unit_count) %>%
    mutate(year = as.integer(year), unit_count = as.numeric(unit_count))
  
  bornes <- long %>%
    group_by(unit_name) %>%
    summarise(debut = min(year), fin = max(year), .groups = "drop")
  
  long %>%
    complete(unit_name, year = annees_obs) %>%
    filter(year %in% annees_obs) %>%
    left_join(bornes, by = "unit_name") %>%
    arrange(unit_name, year)
}

# Même mise en forme finale (large + ancrages 1900/2100) pour les deux versions.
mettre_en_large <- function(long) {
  long %>%
    ungroup() %>%
    pivot_wider(id_cols = unit_name, names_from = year, values_from = unit_count) %>%
    mutate(!!as.character(ANCRAGE_BAS)  := 0,
           !!as.character(ANCRAGE_HAUT) := 0) %>%
    select(unit_name, all_of(as.character(ANNEES_COL)))
}

# Version brute : ni interpolation ni comblement des trous, seulement
# l'ancrage à 0 hors de la période de service. Alimente l'onglet Stocks.
preparer_famille_brut <- function(df, annees_obs = ANNEES_OBS) {
  completer_grille(df, annees_obs) %>%
    group_by(unit_name) %>%
    mutate(
      unit_count = if_else(year >= debut & year <= fin, unit_count, 0),
      unit_count = replace_na(unit_count, 0)
    ) %>%
    mettre_en_large()
}

# Version lissée : interpolation linéaire des trous (années absentes, NA de
# corrections manuelles) à l'intérieur de la période de service. Alimente
# l'onglet Stocks_smooth.
preparer_famille_lisse <- function(df, annees_obs = ANNEES_OBS) {
  completer_grille(df, annees_obs) %>%
    group_by(unit_name) %>%
    mutate(
      unit_count = if (sum(!is.na(unit_count)) >= 2)
        na.approx(unit_count, x = year, na.rm = FALSE, rule = 2) else unit_count,
      unit_count = if_else(year >= debut & year <= fin, unit_count, 0),
      unit_count = replace_na(unit_count, 0)
    ) %>%
    mettre_en_large()
}


### 3. Écriture du classeur ### ----

# familles : liste nommée de data.frames longs, noms = noms de feuilles
ecrire_stocks_xlsx <- function(familles,
                               fichier = "StocksMilitaryMods.xlsx",
                               prefixes = PREFIXES) {
  
  stopifnot(all(names(familles) %in% names(prefixes)))
  
  larges       <- map(familles, preparer_famille_brut)
  larges_lisse <- map(familles, preparer_famille_lisse)
  
  # code Vensim | famille | nom d'usage | années, pour un jeu de feuilles large donné
  assembler_stocks <- function(larges) {
    imap_dfr(larges, function(tb, feuille) {
      tibble(
        code    = paste0(prefixes[[feuille]], "_", nom_vensim(tb$unit_name)),
        famille = prefixes[[feuille]],
        nom     = tb$unit_name
      ) %>% bind_cols(tb %>% select(-unit_name))
    })
  }
  
  stocks       <- assembler_stocks(larges)
  stocks_lisse <- assembler_stocks(larges_lisse)
  
  wb <- createWorkbook()
  
  ## -- Stocks (brut, A1:C1 vides, années à partir de D1) --
  addWorksheet(wb, "Stocks")
  writeData(wb, "Stocks", t(ANNEES_COL), startRow = 1, startCol = 4, colNames = FALSE)
  writeData(wb, "Stocks", stocks,        startRow = 2, startCol = 1, colNames = FALSE)
  
  ## -- Stocks_smooth (lissé, même format) --
  addWorksheet(wb, "Stocks_smooth")
  writeData(wb, "Stocks_smooth", t(ANNEES_COL), startRow = 1, startCol = 4, colNames = FALSE)
  writeData(wb, "Stocks_smooth", stocks_lisse,  startRow = 2, startCol = 1, colNames = FALSE)
  
  ## -- une feuille de contrôle par famille (valeurs brutes) --
  for (feuille in names(larges)) {
    tb <- larges[[feuille]]
    addWorksheet(wb, feuille)
    writeData(wb, feuille, "unit_name", startRow = 1, startCol = 1, colNames = FALSE)
    writeData(wb, feuille, t(ANNEES_COL), startRow = 1, startCol = 2, colNames = FALSE)
    writeData(wb, feuille, tb, startRow = 2, startCol = 1, colNames = FALSE)
  }
  
  saveWorkbook(wb, fichier, overwrite = TRUE)
  message("Écrit : ", normalizePath(fichier), " — ", nrow(stocks), " matériels")
  invisible(list(stocks = stocks, stocks_smooth = stocks_lisse))
}


### 4. Appel ### ----

ecrire_stocks_xlsx(
  familles = list(
    Armoured    = tanks_final,
    Aircrafts   = avions_final,
    Helicopters = helico_final,
    Patrol      = patrol_final,
    Surface     = principal_final,
    Submarines  = sub_final
  ),
  fichier = "StocksMilitaryMods.xlsx"
)

#check rapide
lire_stocks <- function(feuille) {
  read_excel("StocksMilitaryMods.xlsx", sheet = feuille, col_names = FALSE) %>%
    set_names(c("code", "famille", "nom", as.character(unlist(.[1, 4:ncol(.)])))) %>%
    slice(-1) %>%
    pivot_longer(-(code:nom), names_to = "year", values_to = "unit_count") %>%
    mutate(year = as.integer(year), unit_count = as.numeric(unit_count)) %>%
    filter(year >= 1970, year <= 2014)
}

stocks <- lire_stocks("Stocks_smooth")   # ou "Stocks_smooth"

# une famille à la fois (ex. MilitarySurface)
stocks %>%
  # filter(famille == "MilitarySurface") %>%
  ggplot(aes(x = year, y = unit_count, fill = nom)) +
  geom_area() +
  theme_bw() +
  theme(legend.position = "none", legend.text = element_text(size = 6)) +
  guides(fill = guide_legend(ncol = 6, byrow = TRUE))+
  labs(title=paste0("Onglet Stocks_smooth")) # ou "Stocks_smooth"

### Masse totale = stocks lissés × masse unitaire ### ----

masses <- read_excel("military_mass.xlsx", sheet = "Military_MassUnit") %>%
  select(code, mass_t)

stocks_mass <- stocks %>%
  left_join(masses, by = "code") %>%
  mutate(mass_total_t = unit_count * mass_t)

# check rapide : matériels du fichier stocks sans masse associée
stocks_mass %>%
  filter(is.na(mass_t)) %>%
  distinct(code, famille, nom)

# masse totale par famille et par année
masse_par_famille <- stocks_mass %>%
  group_by(famille, year) %>%
  summarise(mass_total_t = sum(mass_total_t, na.rm = TRUE), .groups = "drop")

# graphique 
masse_par_famille %>%
  ggplot(aes(x = year, y = mass_total_t, fill = famille)) +
  geom_area() +
  theme_bw() +
  labs(title = "Masse totale des matériels militaires (stocks lissés)",
       y = "Masse (t)", x = NULL)

# masse totale par unit et par année
masse_par_unit <- stocks_mass %>%
  group_by(nom, year) %>%
  summarise(mass_total_t = sum(mass_total_t, na.rm = TRUE), .groups = "drop") 

# graphique 
masse_par_unit %>%
  filter(mass_total_t > 10*mean(mass_total_t, na.rm = TRUE))|> 
  ggplot(aes(x = year, y = mass_total_t, fill = nom)) +
  geom_area() +
  theme_bw() +
  labs(title = "Plus gros contributeurs au parc en unit_count x masse (10*mean)",
       y = "Masse (t)", x = NULL)
