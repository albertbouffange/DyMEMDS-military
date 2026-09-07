library(tidyverse)
# library(viridis)
# library(ggpubr)
# library(hrbrthemes)
library(zoo)
library(readxl)
library(openxlsx)

# setwd("/home/abouffan/Nextcloud/thèse/dimends/Vensim/")
# dymemds=read_delim("v7.tab",delim = "\t",trim_ws=TRUE) |>
#   filter(if_any(1, ~ str_detect(.x, "Military|Time"))) |>
#   mutate(`2100`=as.double(`2100`)) |>
#   pivot_longer(cols=-1,
#                names_to="Year") |>
#   separate_wider_delim(
#     cols = 1,
#     delim = "[",
#     names = c("Variable", "Modalités"),
#     too_many = "merge",   # si plusieurs "[", on ne coupe qu'au premier
#     too_few = "align_start" # lignes sans "[" -> apres = NA
#   ) |>
#   mutate(Year=as.numeric(Year)) |>
  # mutate(value=as.numeric(value)) |> 
# filter(str_detect(Modalités, "France"))
setwd("/home/abouffan/Nextcloud/R/rDMC")
# dymemds |> write_csv("dymemds_mil_R.csv")
dymemds=read_csv("dymemds_mil_R.csv")

unique(dymemds$Variable)

#variations stocks
var="FlowEoLTUMilitaryFam(/yr)"
dymemds |> 
  filter(Variable==var) |> 
  group_by(Year,Variable) |> 
  ggplot(aes(x=Year,y=value,fill=Modalités))+
  geom_area()+
  scale_x_continuous(limits=c(1980,2010))+
  theme_bw()+
  theme(legend.position="bottom",
                legend.text = element_text(size=6),
                legend.key.size = unit(0.2, 'cm'))+
  labs(title=var)

var="FlowNewTUMilitaryUnit(/yr)"
dymemds |> 
  filter(Variable==var) |> 
  group_by(Year,Variable) |> 
  ggplot(aes(x=Year,y=value,fill=Modalités))+
  geom_area()+
  scale_x_continuous(limits=c(1980,2010))+
  theme_bw()+
  theme(legend.position="none")+
  labs(title=var)

#EmbTU
var="EmbTUMilitary(PJ/yr)"
dymemds |> 
  filter(Variable==var) |> 
  group_by(Year,Variable) |> 
  mutate(value_mean=rollmean(value,3,na.pad=TRUE)) |> 
  ggplot(aes(x=Year,y=value_mean,fill=Modalités))+
  geom_area()+
  scale_x_continuous(limits=c(1982,2010))+
  theme_bw()+
  theme(legend.position="none",
        legend.text = element_text(size=10),
        legend.key.size = unit(0.2, 'cm'))+
    guides(fill=guide_legend(ncol=2,byrow=TRUE))+
  labs(title=var)

dymemds |> 
  filter(Variable==var) |> 
  group_by(Year,Variable) |> 
  summarise(value=sum(value)) |> 
  mutate(value_mean=rollmean(value,3,na.pad=TRUE)) |> 
  ggplot(aes(x=Year,y=value))+
  geom_line()+
  scale_x_continuous(limits=c(1982,2010))+
  scale_y_continuous(limits=c(0,30))+
  theme_bw()+
  labs(title=var)

#ERM
dymemds |> 
  filter(Variable=="ERMMilitary(PJ/yr)") |> 
  group_by(Year,Variable) |> 
  ggplot(aes(x=Year,y=value,fill=Modalités))+
  geom_area()+
  scale_x_continuous(limits=c(1980,2010))+
  theme_bw()+
  theme(legend.position="bottom")

#Énergie
dymemds |> 
  filter(Variable%in%c("ERMMilitary(PJ/yr)","EmbTUMilitary(PJ/yr)")) |> 
  group_by(Year,Variable) |> 
  summarise(value=sum(value)) |> 
  # mutate(value_mean=rollmean(value,3,na.pad=TRUE)) |> 
  ggplot(aes(x=Year,y=value,fill=Variable))+
  geom_area()+
  scale_x_continuous(limits=c(1981,2010))+
  theme_bw()+
  theme(legend.position="bottom")

#RM
var="FlowRMMilitaryPR"
dymemds |> 
  filter(Variable==var) |> 
  mutate(value_mean=rollmean(value,3,na.pad=TRUE)) |> 
  group_by(Year,Variable) |> 
  ggplot(aes(x=Year,y=value_mean,fill=Modalités))+
  geom_area()+
  scale_x_continuous(limits=c(1981,2010))+
  theme_bw()+
  theme(legend.position="bottom",
        legend.text = element_text(size=9),
        legend.key.size = unit(0.2, 'cm'))+
  guides(fill=guide_legend(ncol=3,byrow=TRUE))+
  labs(title=var)

var="FlowRMTUPrimaryMilitaryUnit(t/yr)"
dymemds |> 
  filter(Variable==var) |> 
  group_by(Year,Variable) |> 
  summarise(value=sum(value)) |> 
  # mutate(value_mean=rollmean(value,3,na.pad=TRUE)) |> 
  # ggplot(aes(x=Year,y=value_mean))+
  ggplot(aes(x=Year,y=value))+
  geom_line()+
  scale_x_continuous(limits=c(1981,2010))+
  scale_y_continuous(limits=c(0,11000))+
  theme_bw()+
  labs(title=var)


#####################
#étalement 5 ans avant 
etaler_toutes <- function(df, col_temps, col_valeur, n_annees = 5, groupes = NULL) {
  col_temps  <- rlang::sym(col_temps)
  col_valeur <- rlang::sym(col_valeur)
  
  if (!is.null(groupes)) {
    df <- df %>% group_by(across(all_of(groupes)))
  }
  
  df <- df %>% arrange(!!col_temps, .by_group = !is.null(groupes))
  
  contributions <- df %>%
    transmute(t_source = !!col_temps, part = !!col_valeur / n_annees) %>%
    mutate(cible = map(t_source, ~ (.x - n_annees):(.x - 1))) %>%
    unnest(cible) %>%
    group_by(cible, .add = TRUE) %>%
    summarise(ajout = sum(part), .groups = "drop") %>%
    rename(!!col_temps := cible)
  
  by_cols <- c(rlang::as_string(col_temps), groupes)
  
  df %>%
    ungroup() %>%
    left_join(contributions, by = by_cols) %>%
    mutate(!!col_valeur := coalesce(ajout, 0)) %>%
    select(-ajout)
}

dymemds_décalé <- dymemds %>%
  filter(Variable%in%c("FlowRMMilitaryPR","EmbTU(PJ/yr)")) |> 
  filter(Year>=1980 & Year<=2010) |> 
  etaler_toutes(
    col_temps  = "Year",
    col_valeur = "value",
    n_annees   = 5,
    groupes    = c("Variable", "Modalités")
  )
var="FlowRMMilitaryPR"
dymemds_décalé |> 
  filter(Variable==var) |> 
  group_by(Year,Variable) |> 
  ggplot(aes(x=Year,y=value,fill=Modalités))+
  geom_area()+
  scale_x_continuous(limits=c(1980,2010))+
  theme_bw()+
  theme(legend.position="bottom")+
  labs(title=var)

sans_cdg= dymemds |>  # filter FlowRMTUPrimaryMilitaryUnit(t/yr) str detect CharlesDeGaulle 
  # + summarize
  # +décaler
  filter(Year>=1980 & Year<=2010) |> 
  filter(Variable=="FlowRMTUPrimaryMilitaryUnit(t/yr)") |> 
  filter(!str_detect(Modalités, "CharlesDeGaulle")) |> 
  group_by(Year,Variable) |> 
  summarise(value_sans_cdg=sum(value)) |> 
  etaler_toutes(
    col_temps  = "Year",
    col_valeur = "value_sans_cdg",
    n_annees   = 10,
    groupes    = c("Variable")
  )

avec_cdg= dymemds |>  # avec_cdg= summarise FlowRMTUPrimaryMilitaryUnit(t/yr) 
  # + décaler
  filter(Year>=1980 & Year<=2010) |> 
  filter(Variable=="FlowRMTUPrimaryMilitaryUnit(t/yr)") |> 
  group_by(Year,Variable) |> 
  summarise(value_avec_cdg=sum(value)) |> 
  etaler_toutes(
    col_temps  = "Year",
    col_valeur = "value_avec_cdg",
    n_annees   = 10,
    groupes    = c("Variable")
  )
comparaison_cdg=left_join(sans_cdg,avec_cdg)

ggplot(comparaison_cdg, aes(x=Year))+
  geom_line(aes(y=value_avec_cdg))+
  geom_line(aes(y=value_sans_cdg))+
  labs(title="Avec et sans CDG, 5 ans")

ggplot(comparaison_cdg, aes(x=Year))+
  geom_line(aes(y=value_avec_cdg))+
  geom_line(aes(y=value_sans_cdg))+
  labs(title="Avec et sans CDG, 10 ans")

#Correctifs d'ampleur

# dymemds_décalé2 <- dymemds %>%
#   filter(Year>=1980 & Year<=2020) |> 
#   filter(Variable%in%c("FlowRMMilitaryPR","EmbTU(PJ/yr)")) |> 
#   group_by(Variable,Year) |> 
#   summarise(value=sum(value)) |> 
#   etaler_toutes(
#     col_temps  = "Year",
#     col_valeur = "value",
#     n_annees   = 5,
#     groupes    = c("Variable")
#   )

dymemds_décalé2 <- dymemds %>%
  filter(Year>=1980 & Year<=2020) |> 
  filter(Variable=="FlowRMTUPrimaryMilitaryUnit(t/yr)") |> 
  filter(!str_detect(Modalités, "CharlesDeGaulle")) |> 
  group_by(Year,Variable) |> 
  summarise(value=sum(value)) |> 
  etaler_toutes(
    col_temps  = "Year",
    col_valeur = "value",
    n_annees   = 5,
    groupes    = c("Variable")
  )
exp=read_csv("exportations.csv") 
dymemds_final=left_join(dymemds_décalé2,exp)|>
  filter(Year>=1980 & Year<=2008) |> 
  mutate(value_parc=value*1.2) |> 
  mutate(value_prod=value_parc*1/(1-`CA Export / CA Total`))

dymemds_final |>
  ggplot(aes(x=Year))+
  geom_line(aes(y=value,colour="blue"))+
  geom_line(aes(y=value_parc,colour="red"))+
  geom_line(aes(y=value_prod))+
  scale_y_continuous(
    name = "FlowRMMilitaryPR",
    sec.axis = sec_axis(name="CA Export / CA Total")
  )+
  geom_line(aes(y=`CA Export / CA Total`))
theme_bw()+
  theme(legend.position = "none")+
  labs(title="Production décalée de 5 ans : modèle dymemds, total armées, total production armement")
