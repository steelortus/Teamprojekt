# TEAMPROJEKT OOXML-Forensik

Powershell Skripte für OOXML-Forensik

Scripts have to be run in the same folder as the document that is is to be run on.

run './main.ps1'


## Current Tasks:
- [x] Script um OO dokumente in Folder zu konvertieren
- [ ] Word Excel PP where metadata? (Abgesehen davon erstmal nur Word)
- [ ] metadata xml to csv via xslt
- [ ] Abweichende Metadaten von MSOffice LibreOffice OpenOffice
- [ ] Metadaten auswerten in Bezug auf Datierung (Daten Created/Modified/Printed, Revision, Releasetermin Softwareversion, Releasetermine Fonts)
- [ ] In RSIDs einarbeiten
- [ ] Anzahl Zeichen im Dokument ermitteln (mit und ohne Leerzeichen) Ermittelte Anzahl Zeichen im Dokument vergleichen mit zugehörigem Wert in Metadaten

## Other Tasks:
- [ ] dateiendung nicht verwenden um die richtige auslese zu finden
- [ ] datastructure for versions, releases, fontreleases, also: fontrelease vs fontrelease in software
- [ ] rsid (revision safe identifiers) nutzen finden. verleich word vs libreoffice

## Extra:
Folder -> Get oldets doc/ get newest doc


## Thinking for later:
- Was ist versionsnummer? Intern öffentlich identisch? Wann auf den markt gekommen
- Datastructure for changes across versions and versions across time
- change metadata to test calculations vs read (number of words)
