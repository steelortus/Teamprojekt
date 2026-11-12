# TEAMPROJEKT OOXML-Forensik

Powershell Skripte für OOXML-Forensik

Scripts have to be run in the same folder as the document that is is to be run on.

run './main.ps1'


## Current Tasks:
- [x] Script um OO dokumente in Folder zu konvertieren
- [x] Word Excel PP where metadata? (Abgesehen davon erstmal nur Word)
- [x] metadata xml to csv via xslt
- [ ] Abweichende Metadaten von MSOffice LibreOffice OpenOffice
- [ ] Metadaten auswerten in Bezug auf Datierung (Daten Created/Modified/Printed, Revision, Releasetermin Softwareversion, Releasetermine Fonts, Language)
- [x] In RSIDs einarbeiten
- [x] Anzahl Zeichen im Dokument ermitteln (mit und ohne Leerzeichen) Ermittelte Anzahl Zeichen im Dokument vergleichen mit zugehörigem Wert in Metadaten
- [ ] Anzahl Absätze im Dokumentermitteln und vergleichen mit Metadaten
- [ ] Im Dokument gekennzeichnete Sprachen für Textabschnitte ermitteln
- [ ] In welchen Zeiträumen wurden diese Sprachkennzeichnungen verwendet (ab welchem Zeitpunkt eingeführt/veraltet?)
- [ ] Ermittlung RSIDs für Dokument, paragraphs, runs
- [ ] Abgleich RSIDs in Dokument mit RSIDs in Metadaten

## Optional Tasks:
- [x] dateiendung nicht verwenden um die richtige auslese zu finden
- [ ] datastructure for versions, releases, fontreleases, also: fontrelease vs fontrelease in software
- [ ] rsid (revision safe identifiers) nutzen finden. verleich word vs libreoffice
- [ ] change main to scan for ps1 scripts instead of manually saving them.
- [ ] Dev-script to automatically test multiple files
- [ ] What counts as a word (Absätze)
- [ ] What characters are counted by docx that are not counted by script?
- [ ] Compare RSIDs from settings.xml to document.xml
- [ ] Do Word and Writer count differently
- [ ] Change a file multiple times and compare the RSIDs

## Thinking for later:
- [ ] Was ist versionsnummer? Intern öffentlich identisch? Wann auf den markt gekommen
- [x] Datastructure for changes across versions and versions across time -> multiple CSVs
- [x] change metadata to test calculations vs read (number of words)
- [ ] welchen Zeiträumen wurden welche Sprachkennzeichnungen
