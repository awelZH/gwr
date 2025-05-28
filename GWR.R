library(tidyverse)
library(readxl)
library(openxlsx)


Gebaeude_roh <- read.csv2("gebaeude_batiment_edificio.csv", sep = "\t")
Strassen_roh <- read.csv2("eingang_entree_entrata.csv", sep = "\t")
GVZ <- read.csv2("GVZ_VOLUMEN.csv")
Hoehe_GIS <- read.csv2("Gebaeude_hoehe.csv")
load("EBF_Modell.RData")

# erneuerbare definieren
erneuerbare <- c("WP-unb.", "LW-WP", "ES-WP", "EReg-WP", "WW-WP", "Fernw.", "Th.Sol", "Ofen ern.", "Holzkes.")

# Höhe als numerisch
Hoehe_GIS <- Hoehe_GIS %>%
  mutate(HOEHE = as.numeric(HOEHE))

# Codes von GWR zu Text übersetzten
Gebaeude <- Gebaeude_roh %>%
  mutate(
    Heizsystem_1_berechnet = case_when(
      GWAERZH1 %in% c(7410, 7411) ~ case_when( # Wärmepumpen
        GENH1 %in% c(7598, 7500, 7599) ~ "WP-unb.",
        GENH1 == 7501 ~ "LW-WP",
        GENH1 %in% c(7510, 7511) ~ "ES-WP",
        GENH1 == 7512 ~ "EReg-WP",
        GENH1 == 7513 ~ "WW-WP",
        TRUE ~ "WP-unb."
      ),
      GWAERZH1 %in% c(7460, 7461) ~ "Fernw.", # Fernwärme
      
      GWAERZH1 > 7429 & GWAERZH1 < 7436 ~ case_when( # Heizkessel
        GENH1 %in% c(7540, 7541, 7542, 7543) ~ "Holzkes.",
        GENH1 == 7520 ~ "Gaskes.",
        GENH1 == 7530 ~ "Oelkes.",
        TRUE ~ "Kessel unb."
      ),
      
      GWAERZH1 == 7436 ~ case_when( #Ofen
        GENH1 >= 7540 & GENH1 <= 7543 ~ "Ofen ern.",
        GENH1 >= 7520 & GENH1 <= 7530 ~ "Ofen foss.",
        GENH1 == 7560 ~ "El.Ofen",
        TRUE ~ "Ofen unb."
      ), 
      
      # Andere Heizungen
      GWAERZH1 == 7400 ~ "Kein WärErz",
      GWAERZH1 %in% c(7420, 7421) ~ "Th.Sol",
      GWAERZH1 %in% c(7440, 7441) ~ "WKK",
      GWAERZH1 %in% c(7450, 7451, 7452) ~ "El.Heizung",
      GWAERZH1 == 7499 ~ "Andere"
    ),
    
    Heizsystem_2_berechnet = case_when(
      GWAERZH2 %in% c(7410, 7411) ~ case_when( # Wärmepumpen
        GENH2 %in% c(7598, 7500, 7599) ~ "WP-unb.",
        GENH2 == 7501 ~ "LW-WP",
        GENH2 %in% c(7510, 7511) ~ "ES-WP",
        GENH2 == 7512 ~ "EReg-WP",
        GENH2 == 7513 ~ "WW-WP",
        TRUE ~ "WP-unb."
      ),
      GWAERZH2 %in% c(7460, 7461) ~ "Fernw.", # Fernwärme
      
      GWAERZH2 > 7429 & GWAERZH2 < 7436 ~ case_when( # Heizkessel
        GENH2 %in% c(7540, 7541, 7542, 7543) ~ "Holzkes.",
        GENH2 == 7520 ~ "Gaskes.",
        GENH2 == 7530 ~ "Oelkes.",
        TRUE ~ "Kessel unb."
      ),
      
      GWAERZH2 == 7436 ~ case_when( #Ofen
        GENH2 >= 7540 & GENH2 <= 7543 ~ "Ofen ern.",
        GENH2 >= 7520 & GENH2 <= 7530 ~ "Ofen foss.",
        GENH2 == 7560 ~ "El.Ofen",
        TRUE ~ "Ofen unb."
      ),  
      
      # Andere Heizungen
      GWAERZH2 == 7400 ~ "Kein WärErz",
      GWAERZH2 %in% c(7420, 7421) ~ "Th.Sol",
      GWAERZH2 %in% c(7440, 7441) ~ "WKK",
      GWAERZH2 %in% c(7450, 7451, 7452) ~ "El.Heizung",
      GWAERZH2 == 7499 ~ "Andere"
    ),
    
    Datenquelle_Heizung_1 = case_when(
      GWAERSCEH1 == 852 ~ "Gemäss amtlicher Schätzung",
      GWAERSCEH1 == 853 ~ "Gemäss Gebäudeversicherung",
      GWAERSCEH1 == 855 ~ "Gemäss Feuerkontrolle",
      GWAERSCEH1 == 857 ~ "Eigentümer/in / Verwaltung",
      GWAERSCEH1 == 858 ~ "Gemäss Gebäudeenergieausweis der Kantone (GEAK)",
      GWAERSCEH1 == 859 ~ "Andere",
      GWAERSCEH1 == 860 ~ "Gemäss Volkszählung 2000",
      GWAERSCEH1 == 864 ~ "Gemäss kantonalen Daten / Das Gebäudeprogramm",
      GWAERSCEH1 == 865 ~ "Gemäss Daten der Gemeinde",
      GWAERSCEH1 == 869 ~ "Gemäss Baubewilligung",
      GWAERSCEH1 == 870 ~ "Gemäss Versorgungswerk (Gas, Fernwärme)",
      GWAERSCEH1 == 871 ~ "Gemäss Minergie"
    ),
    
    Datenquelle_Heizung_2 = case_when(
      GWAERSCEH2 == 852 ~ "Gemäss amtlicher Schätzung",
      GWAERSCEH2 == 853 ~ "Gemäss Gebäudeversicherung",
      GWAERSCEH2 == 855 ~ "Gemäss Feuerkontrolle",
      GWAERSCEH2 == 857 ~ "Eigentümer/in / Verwaltung",
      GWAERSCEH2 == 858 ~ "Gemäss Gebäudeenergieausweis der Kantone (GEAK)",
      GWAERSCEH2 == 859 ~ "Andere",
      GWAERSCEH2 == 860 ~ "Gemäss Volkszählung 2000",
      GWAERSCEH2 == 864 ~ "Gemäss kantonalen Daten / Das Gebäudeprogramm",
      GWAERSCEH2 == 865 ~ "Gemäss Daten der Gemeinde",
      GWAERSCEH2 == 869 ~ "Gemäss Baubewilligung",
      GWAERSCEH2 == 870 ~ "Gemäss Versorgungswerk (Gas, Fernwärme)",
      GWAERSCEH2 == 871 ~ "Gemäss Minergie"
    ),
    
    Wassererwärmer_1_berechnet = case_when(
      GWAERZW1 %in% c(7610) ~ case_when( # Wärmepumpen
        GENW1 %in% c(7598, 7500, 7599) ~ "WP-unb.",
        GENW1 == 7501 ~ "LW-WP",
        GENW1 %in% c(7510, 7511) ~ "ES-WP",
        GENW1 == 7512 ~ "EReg-WP",
        GENW1 == 7513 ~ "WW-WP",
        TRUE ~ "WP-unb."
      ),
      GWAERZW1 %in% c(7660) ~ "Fernw.", # Fernwärme
      
      GWAERZW1 > 7629 & GWAERZW1 < 7635 ~ case_when( # Heizkessel
        GENW1 %in% c(7540, 7541, 7542, 7543) ~ "Holzkes.",
        GENW1 == 7520 ~ "Gaskes.",
        GENW1 == 7530 ~ "Oelkes.",
        TRUE ~ "Kessel unb."
      ),
      # Andere Warmwassererzeuger
      GWAERZW1 == 7600 ~ "Kein WärErz",
      GWAERZW1 %in% c(7620) ~ "Th.Sol",
      GWAERZW1 %in% c(7640) ~ "WKK",
      GWAERZW1 %in% c(7651) ~ "Kl.Boil",
      GWAERZW1 %in% c(7650) ~ "Zent.El.Boil",
      GWAERZW1 == 7699 ~ "Andere"
    ),
    
    Wassererwärmer_2_berechnet = case_when(
      GWAERZW2 %in% c(7610) ~ case_when( # Wärmepumpen
        GENW2 %in% c(7598, 7500, 7599) ~ "WP-unb.",
        GENW2 == 7501 ~ "LW-WP",
        GENW2 %in% c(7510, 7511) ~ "ES-WP",
        GENW2 == 7512 ~ "EReg-WP",
        GENW2 == 7513 ~ "WW-WP",
        TRUE ~ "WP-unb."
      ),
      GWAERZW2 %in% c(7660) ~ "Fernw.", # Fernwärme
      
      GWAERZW2 > 7629 & GWAERZW2 < 7635 ~ case_when( # Heizkessel
        GENW2 %in% c(7540, 7541, 7542, 7543) ~ "Holzkes.",
        GENW2 == 7520 ~ "Gaskes.",
        GENW2 == 7530 ~ "Oelkes.",
        TRUE ~ "Kessel unb."
      ),
      # Andere Warmwassererzeuger
      GWAERZW2 == 7600 ~ "Kein WärErz",
      GWAERZW2 %in% c(7620) ~ "Th.Sol",
      GWAERZW2 %in% c(7640) ~ "WKK",
      GWAERZW2 %in% c(7651) ~ "Kl.Boil",
      GWAERZW2 %in% c(7650) ~ "Zent.El.Boil",
      GWAERZW2 == 7699 ~ "Andere"
    ), 
    
    Datenquelle_Wassererwärmung_1 = case_when(
      GWAERSCEW1 == 852 ~ "Gemäss amtlicher Schätzung",
      GWAERSCEW1 == 853 ~ "Gemäss Gebäudeversicherung",
      GWAERSCEW1 == 855 ~ "Gemäss Feuerkontrolle",
      GWAERSCEW1 == 857 ~ "Eigentümer/in / Verwaltung",
      GWAERSCEW1 == 858 ~ "Gemäss Gebäudeenergieausweis der Kantone (GEAK)",
      GWAERSCEW1 == 859 ~ "Andere",
      GWAERSCEW1 == 860 ~ "Gemäss Volkszählung 2000",
      GWAERSCEW1 == 864 ~ "Gemäss kantonalen Daten / Das Gebäudeprogramm",
      GWAERSCEW1 == 865 ~ "Gemäss Daten der Gemeinde",
      GWAERSCEW1 == 869 ~ "Gemäss Baubewilligung",
      GWAERSCEW1 == 870 ~ "Gemäss Versorgungswerk (Gas, Fernwärme)",
      GWAERSCEW1 == 871 ~ "Gemäss Minergie"
    ),
    
    Datenquelle_Wassererwärmung_2 = case_when(
      GWAERSCEW2 == 852 ~ "Gemäss amtlicher Schätzung",
      GWAERSCEW2 == 853 ~ "Gemäss Gebäudeversicherung",
      GWAERSCEW2 == 855 ~ "Gemäss Feuerkontrolle",
      GWAERSCEW2 == 857 ~ "Eigentümer/in / Verwaltung",
      GWAERSCEW2 == 858 ~ "Gemäss Gebäudeenergieausweis der Kantone (GEAK)",
      GWAERSCEW2 == 859 ~ "Andere",
      GWAERSCEW2 == 860 ~ "Gemäss Volkszählung 2000",
      GWAERSCEW2 == 864 ~ "Gemäss kantonalen Daten / Das Gebäudeprogramm",
      GWAERSCEW2 == 865 ~ "Gemäss Daten der Gemeinde",
      GWAERSCEW2 == 869 ~ "Gemäss Baubewilligung",
      GWAERSCEW2 == 870 ~ "Gemäss Versorgungswerk (Gas, Fernwärme)",
      GWAERSCEW2 == 871 ~ "Gemäss Minergie"
    ),
    
    Mehrere_Gebaeude = case_when(
      GWAERZH1 %in% c(7411, 7421, 7431, 7435, 7441, 7451, 7461) ~ "Ja",
      GWAERZH2 %in% c(7411, 7421, 7431, 7435, 7441, 7451, 7461) ~ "Ja",
      TRUE ~ "Nein"
    ),
    
    Datum_Heizung_1 = as.Date(GWAERDATH1, format = "%Y-%m-%d"),
    Datum_Heizung_2 = as.Date(GWAERDATH2, format = "%Y-%m-%d"),
    Datum_Warwassererzeuger_1 = as.Date(GWAERDATW1, format = "%Y-%m-%d"),
    Datum_Warwassererzeuger_2 = as.Date(GWAERDATW2, format = "%Y-%m-%d"),
    
    # Excel-kompatible Zahlenformate erstellen (für Datum)
    Excel_Datum_Heizung_1 = as.numeric(Datum_Heizung_1) + 25569,
    Excel_Datum_Heizung_2 = as.numeric(Datum_Heizung_2) + 25569,
    Excel_Datum_Warwassererzeuger_1 = as.numeric(Datum_Warwassererzeuger_1) + 25569,
    Excel_Datum_Warwassererzeuger_2 = as.numeric(Datum_Warwassererzeuger_2) + 25569,
    
    # Alternative Darstellung als "dd.mm.yyyy" (Character-Format)
    Format_Datum_Heizung_1 = format(Datum_Heizung_1, "%d.%m.%Y"),
    Format_Datum_Heizung_2 = format(Datum_Heizung_2, "%d.%m.%Y"),
    Format_Datum_Warwassererzeuger_1 = format(Datum_Warwassererzeuger_1, "%d.%m.%Y"),
    Format_Datum_Warwassererzeuger_2 = format(Datum_Warwassererzeuger_2, "%d.%m.%Y"),
    
    # Prüfen ob fossile Systeme nach 01.09.2022 bewilligt wurden
    H1_EnerG = case_when(
      GWAERSCEH1 == 869 & as.Date(Datum_Heizung_1, format = "%d.%m.%Y") >= as.Date("01.09.2022", format = "%d.%m.%Y") ~ Datum_Heizung_1
    ),
    H2_EnerG = case_when(
      GWAERSCEH2 == 869 & as.Date(Datum_Heizung_2, format = "%d.%m.%Y") >= as.Date("01.09.2022", format = "%d.%m.%Y") ~ Datum_Heizung_2
    ),
    W1_EnerG = case_when(
      GWAERSCEW1 == 869 & as.Date(Datum_Warwassererzeuger_1, format = "%d.%m.%Y") >= as.Date("01.09.2022", format = "%d.%m.%Y") ~ Datum_Warwassererzeuger_1
    ),
    W2_EnerG = case_when(
      GWAERSCEW2 == 869 & as.Date(Datum_Warwassererzeuger_2, format = "%d.%m.%Y") >= as.Date("01.09.2022", format = "%d.%m.%Y") ~ Datum_Warwassererzeuger_2
    ),
    
    # Fall: Sowohl erneuerbare als auch fossile Systeme sind vorhanden
    System = case_when(
      ((H1_EnerG > 0 & str_detect(Heizsystem_1_berechnet, paste(erneuerbare, collapse = "|"))) |
         (H2_EnerG > 0 & str_detect(Heizsystem_2_berechnet, paste(erneuerbare, collapse = "|"))) |
         (W1_EnerG > 0 & str_detect(Wassererwärmer_1_berechnet, paste(erneuerbare, collapse = "|"))) |
         (W2_EnerG > 0 & str_detect(Wassererwärmer_2_berechnet, paste(erneuerbare, collapse = "|")))) &
        ((H1_EnerG > 0 & Heizsystem_1_berechnet != "Kein WärErz" & !str_detect(Heizsystem_1_berechnet, paste(erneuerbare, collapse = "|"))) |
           (H2_EnerG > 0 & Heizsystem_2_berechnet != "Kein WärErz" & !str_detect(Heizsystem_2_berechnet, paste(erneuerbare, collapse = "|"))) |
           (W1_EnerG > 0 & Wassererwärmer_1_berechnet != "Kein WärErz" & !str_detect(Wassererwärmer_1_berechnet, paste(erneuerbare, collapse = "|"))) |
           (W2_EnerG > 0 & Wassererwärmer_2_berechnet != "Kein WärErz" & !str_detect(Wassererwärmer_2_berechnet, paste(erneuerbare, collapse = "|")))) ~ "Ern.&Foss.",
      
      
      # Fall 1: EnerG > 0, erneuerbares System
      (H1_EnerG > 0 & str_detect(Heizsystem_1_berechnet, paste(erneuerbare, collapse = "|"))) |
        (H2_EnerG > 0 & str_detect(Heizsystem_2_berechnet, paste(erneuerbare, collapse = "|"))) |
        (W1_EnerG > 0 & str_detect(Wassererwärmer_1_berechnet, paste(erneuerbare, collapse = "|"))) |
        (W2_EnerG > 0 & str_detect(Wassererwärmer_2_berechnet, paste(erneuerbare, collapse = "|"))) ~ "Ern.",
      
      
      # Fall 2: EnerG > 0, aber kein erneuerbares System
      (H1_EnerG > 0 & Heizsystem_1_berechnet != "Kein WärErz") |
        (H2_EnerG > 0 & Heizsystem_2_berechnet != "Kein WärErz") |
        (W1_EnerG > 0 & Wassererwärmer_1_berechnet != "Kein WärErz") |
        (W2_EnerG > 0 & Wassererwärmer_2_berechnet != "Kein WärErz") ~ "Foss.",
      
      # Fall 3: Kein Wärmeerzeuger
      (H1_EnerG > 0 & Heizsystem_1_berechnet == "Kein WärErz") |
        (H2_EnerG > 0 & Heizsystem_2_berechnet == "Kein WärErz") |
        (W1_EnerG > 0 & Wassererwärmer_1_berechnet == "Kein WärErz") |
        (W2_EnerG > 0 & Wassererwärmer_2_berechnet == "Kein WärErz") ~ "Kein WärErz",
      
      # Standardfall: NA, wenn keine der Bedingungen zutrifft
      (H1_EnerG > 0 | H2_EnerG > 0 | W1_EnerG > 0 | W2_EnerG > 0)  ~ "Unbekannt"
    ),
    
    # 8-Jahre oder mehr unterschied zwischen Genehmigung und Baujahr Gebäude = Ersatz. Unbestimmt bedeutet, dass nicht klar ist ob es ein Ersatz ist.
    # Zum Beispiel wenn der Ersatz und das Baujahr weniger als 8 Jahre auseinander liegen.
    Ersatz = case_when(
      (as.integer(format(H1_EnerG, "%Y"))) > 0 & ((as.integer(format(H1_EnerG, "%Y"))) - 8) > GBAUJ & System != "Kein WärErz" ~ "Ersatz.H1",
      (as.integer(format(H1_EnerG, "%Y"))) > 0 & ((as.integer(format(H1_EnerG, "%Y"))) - 8) <= GBAUJ & System != "Kein WärErz" ~ "Unbes.H1",
      (as.integer(format(H2_EnerG, "%Y"))) > 0 & ((as.integer(format(H2_EnerG, "%Y"))) - 8) > GBAUJ & System != "Kein WärErz" ~ "Ersatz.H2",
      (as.integer(format(H2_EnerG, "%Y"))) > 0 & ((as.integer(format(H2_EnerG, "%Y"))) - 8) <= GBAUJ & System != "Kein WärErz" ~ "Unbes.H2",
      (as.integer(format(W1_EnerG, "%Y"))) > 0 & ((as.integer(format(W1_EnerG, "%Y"))) - 8) > GBAUJ & System != "Kein WärErz" ~ "Ersatz.W1",
      (as.integer(format(W1_EnerG, "%Y"))) > 0 & ((as.integer(format(W1_EnerG, "%Y"))) - 8) <= GBAUJ & System != "Kein WärErz" ~ "Unbes.W1",
      (as.integer(format(W2_EnerG, "%Y"))) > 0 & ((as.integer(format(W2_EnerG, "%Y"))) - 8) > GBAUJ & System != "Kein WärErz" ~ "Ersatz.W2",
      (as.integer(format(W2_EnerG, "%Y"))) > 0 & ((as.integer(format(W2_EnerG, "%Y"))) - 8) <= GBAUJ & System != "Kein WärErz" ~ "Unbes.W2",
      (System == "Kein WärErz") ~ "Kein WärErz",
      is.na(GBAUJ) & (
        (as.integer(format(H1_EnerG, "%Y"))) > 0 |
          (as.integer(format(H2_EnerG, "%Y"))) > 0 |
          (as.integer(format(W1_EnerG, "%Y"))) > 0 |
          (as.integer(format(W2_EnerG, "%Y"))) > 0
      ) ~ "Kein GBAUJ" # Nur wenn GBAUJ NA ist und eine andere Bedingung theoretisch erfüllt wäre
    ),
    
    #Gebäudestatus definieren, nur bestehende interessiern uns.
    GebStatus = case_when(
      GSTAT == 1001 | GSTAT == 1002 | GSTAT == 1003 ~ "Kommt",
      GSTAT == 1004 ~ "Bestehend",
      TRUE ~ "Andere"
    )
  )


# GVZ hinzufügen
Gebaeude <- Gebaeude %>%
  left_join(GVZ, by = "EGID")

# Hoehe hinzufügen
Gebaeude <- Gebaeude %>%
  left_join(Hoehe_GIS, by = "EGID")

# Strassen hinzufügen
Strassen <- Strassen_roh %>%
  mutate(
    Strasse = paste(STRNAME, DEINR, sep = " ")) %>%
  group_by(EGID) %>% # Nur die erste Strasse pro EGID behalten
  slice(1) %>% 
  ungroup()


# Datensätzeverbinden
Gebaeude <- Gebaeude %>%
  left_join(Strassen %>% 
              select(EGID, Strasse, DPLZ4, DPLZZ), by = "EGID")

# EBF prediction
Gebaeude <- Gebaeude %>%
  mutate(
    GAREA  = ifelse(is.na(GAREA), VOLUMEN / HOEHE, GAREA),
    GASTW  = ifelse(is.na(GASTW), HOEHE / 3.5, GASTW),
    HOEHE  = ifelse(is.na(HOEHE), VOLUMEN / GAREA, HOEHE),
    VOLUMEN = ifelse(is.na(VOLUMEN), GAREA * HOEHE, VOLUMEN)
  ) %>%
  mutate(
    log_GAREA = log(GAREA),
    log_GASTW = log(GASTW),
    log_HOEHE = log(HOEHE),
    log_VOLUMEN = log(VOLUMEN)
  ) %>%
  mutate(
    EBF_PREDICTED_LOG = predict(fit_robust, newdata = .),
    EBF_PREDICTED = round(exp(EBF_PREDICTED_LOG))
  ) %>%
  mutate(
    EBF = case_when(
      GEBF >= 2000 ~ GEBF,
      (GEBF >= EBF_PREDICTED * 0.8 & GEBF <= EBF_PREDICTED) | (GEBF <= EBF_PREDICTED * 1.2 & GEBF >= EBF_PREDICTED) ~ GEBF,
      TRUE ~ EBF_PREDICTED
    ),
    FLAG_EBF = case_when( # EBF Flag, wenn 0, dann der bereits vorhadnene EBF-Wert, wenn 1, dann der Predicted-Wert
      EBF == GEBF ~ 0,
      EBF == EBF_PREDICTED ~ 1,
      TRUE ~ NA_real_  # Falls es unerwartete Werte gibt
    )
  )


# Excel erstellen und exportieren
## Reihenfolge der Spalte bestimmen
Gebaeude_export <- Gebaeude %>%
  select(EGID, GGDENAME, Strasse, DPLZ4, DPLZZ,
         GSTAT, GebStatus, GKAT, GBAUJ, GBAUP, 
         EBF, FLAG_EBF, GAREA, GASTW,
         GWAERZH1, GENH1, Heizsystem_1_berechnet,
         Datum_Heizung_1, GWAERSCEH1, Datenquelle_Heizung_1,
         GWAERZH2, GENH2, Heizsystem_2_berechnet,
         Datum_Heizung_2, GWAERSCEH2, Datenquelle_Heizung_2,
         GWAERZW1, GENW1, Wassererwärmer_1_berechnet,
         Datum_Warwassererzeuger_1, GWAERSCEW1, Datenquelle_Wassererwärmung_1,
         GWAERZW2, GENW2, Wassererwärmer_2_berechnet,
         Datum_Warwassererzeuger_2, GWAERSCEW2, Datenquelle_Wassererwärmung_2,
         H1_EnerG, H2_EnerG, W1_EnerG, W2_EnerG, System, Ersatz, Mehrere_Gebaeude
         )


# Duplicate entfernen
Gebaeude_export <- Gebaeude_export %>%
  distinct(EGID, .keep_all = TRUE)

# Excle erstellen
##Spatlen einfärben
wb <- createWorkbook()
addWorksheet(wb, "Sheet1")
writeData(wb, "Sheet1", Gebaeude_export, startRow = 1)

# Styles erstellen
style1 <- createStyle(fgFill = "black", fontColour = "white", textDecoration = "bold", halign = "CENTER")
style2 <- createStyle(fgFill = "lightskyblue", fontColour = "black", textDecoration = "bold", halign = "CENTER")
style3 <- createStyle(fgFill = "salmon2", fontColour = "black", textDecoration = "bold", halign = "CENTER") 
style4 <- createStyle(fgFill = "deepskyblue2", fontColour = "black", textDecoration = "bold", halign = "CENTER")
style5 <- createStyle(fgFill = "coral2", fontColour = "black", textDecoration = "bold", halign = "CENTER") 
style6 <- createStyle(fgFill = "grey23", fontColour = "white", textDecoration = "bold", halign = "CENTER")
style7 <- createStyle(fgFill = "olivedrab", fontColour = "black", textDecoration = "bold", halign = "CENTER")
style8 <- createStyle(fgFill = "goldenrod2", fontColour = "black", textDecoration = "bold", halign = "CENTER")
dateStyle <- createStyle(numFmt = "DATE")

#Styles auf die entsprechenden Spalten anwenden
addStyle(wb, "Sheet1", style1, rows = 1, cols = 1)
addStyle(wb, "Sheet1", style7, rows = 1, cols = 2:5)
addStyle(wb, "Sheet1", style8, rows = 1, cols = 6:14)
addStyle(wb, "Sheet1", style2, rows = 1, cols = 15:20)
addStyle(wb, "Sheet1", style3, rows = 1, cols = 21:26)
addStyle(wb, "Sheet1", style4, rows = 1, cols = 27:32)
addStyle(wb, "Sheet1", style5, rows = 1, cols = 33:38)
addStyle(wb, "Sheet1", style6, rows = 1, cols = 39:45)

# Filter setzten für alle Spalten
addFilter(wb, "Sheet1", rows = 1, cols = 1:45)

# Erste Zeile fixieren
freezePane(wb, "Sheet1" , firstRow = TRUE, firstCol = FALSE)

#Spaltenbreite anpassen
setColWidths(wb, "Sheet1", cols = 1:38, widths = "auto")
setColWidths(wb, "Sheet1", cols = 39:42, widths = 15)
setColWidths(wb, "Sheet1", cols = 43:45, widths = "auto")


# export als xlsx-File
saveWorkbook(wb, "GWR_Gebaeude_bereinigt.xlsx", overwrite = TRUE)







