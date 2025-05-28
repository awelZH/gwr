# R Script zur GWR Auswertung

GWR.R erlaubt es, eine Auswertung der GWR-Daten vorzunehmen. Zusätzlich wird eine EBF-Schätzung für alle Gebäude vorgenommen.<br /> 
Das Schätzungsmodell ist in EBF_Modell.RData enthalten und die Berechnungen dazu sind in meinem Praktikumsordner zu finden.<br /> 
Für die Schätzung, werden einige extrene Daten benötigt, die direkt im RScript eingebunden werden.<br />  

# Anwendung
Damit das Script funktioniert, müssen folgende Daten beim ausführen in der gleichen Directory wie GWR.R sein.<br /> 
Aus den Rohdaten:<br /> 
        - eingang_entree_entrata.csv<br /> 
        - gebaeude_batiment_edificio.csv<br /> 

Externe Daten für EBF-Schätzung:<br /> 
        - GVZ_VOLUMEN.csv (vertrauliche Daten)<br /> 
        - Gebaeude_hoehe.csv (von Gebäude 3D Datensatz der Landestopo)<br /> 
        - EBF_Modell.RData<br /> 

Die Auswertung kann mit GWR_R_SCRIPT.bat gestartet werden, sofern es auch in der selben Directory ist.<br /> 
Tidyverse muss installiert sein(!), ansonsten funktioniert die Auswertung nicht.<br /> 

# Datenbezug
Die aktuellen GWR-Daten können mit GWR_DL_easy.bat per PowerShell heruntergeladen werden. Vor dem Ausführen muss die .bat-Datei bearbeitet werden und der Outputpath angepasst werden.
Dazu kann einfach der gewünschte Path nach -Outfile eingefugt werden (Achtung Anführungszeichen werden benötigt).

# Python
Es gibt eine Möglichkeit die Auswertung zu automatisieren. Leider konnte ich diese Dateinen hier nicht hochladen. Auf dem K-Laufwerk der Energie, in der Directory 1GWR_Script_Master, befinden sich die Python-Files und ein ReadMe mit der Installationsanleitung.
