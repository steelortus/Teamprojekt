# TEAMPROJEKT OOXML-Forensik

Powershell Skripte für OOXML-Forensik

Scripts have to be run in the same folder as the document that is is to be run on.

run './main.ps1'


## Current Tasks:
- [x] Script um OO dokumente in Folder zu konvertieren
- [x] Word Excel PP where metadata? (Abgesehen davon erstmal nur Word)
- [x] metadata xml to csv via xslt
- [x] Abweichende Metadaten von MSOffice LibreOffice OpenOffice
- [ ] Metadaten auswerten in Bezug auf Datierung (Daten Created/Modified/Printed, Revision, Releasetermin Softwareversion, Releasetermine Fonts, Language)
- [x] In RSIDs einarbeiten
- [x] Anzahl Zeichen im Dokument ermitteln (mit und ohne Leerzeichen) Ermittelte Anzahl Zeichen im Dokument vergleichen mit zugehörigem Wert in Metadaten
- [x] Anzahl Absätze im Dokumentermitteln und vergleichen mit Metadaten
- [x] Im Dokument gekennzeichnete Sprachen für Textabschnitte ermitteln
- [x] In welchen Zeiträumen wurden diese Sprachkennzeichnungen verwendet (ab welchem Zeitpunkt eingeführt/veraltet?)
- [ ] Ermittlung RSIDs für Dokument, paragraphs, runs
- [x] Abgleich RSIDs in Dokument mit RSIDs in Metadaten (List all RSIDs only on one of them both)


## Optional Tasks:
- [x] dateiendung nicht verwenden um die richtige auslese zu finden
- [ ] datastructure for versions, releases, fontreleases, also: fontrelease vs fontrelease in software
- [ ] rsid (revision safe identifiers) nutzen finden. verleich word vs libreoffice
- [x] change main to scan for ps1 scripts instead of manually saving them.
- [x] Dev-script to automatically test multiple files
- [?] What counts as a word (Absätze)
- [?] What characters are counted by docx that are not counted by script?
- [x] Compare RSIDs from settings.xml to document.xml
- [x] Do Word and Writer count differently - no
- [ ] Change a file multiple times and compare the RSIDs
- [ ] Get list of languages and percentages how many percent are that language
- [ ] Where are all the RSIDs
- [ ] Are all RSIDs (p/r/etc) chacked in the compare RSIDs script
- [ ] Wie ändern sich RSIDs, wenn Dokumente verändert werden. (neues zeichen, zeichen löschen, neuer paragraph, etc)


## Thinking for later:
- [ ] Was ist versionsnummer? Intern öffentlich identisch? Wann auf den markt gekommen
- [x] Datastructure for changes across versions and versions across time -> multiple CSVs
- [x] change metadata to test calculations vs read (number of words)
- [ ] dev-script to auto tell all RSIDs with one click
- [ ] script um die datei wieder zu zippen




## Issues:
- [ ] No Idea, how things are being counted. Absätze sollten <w:p> sein aber sind nicht.
    - `test_long.docx` 133/35
    - `ProccessorAgreement.docx` 579/58
    - `SoftwareOrder.docx` 43/1
- [ ] RSID sind zu unterschiedlich `ProccessorAgreement.docx`: 170 settings.xml vs 48 document parts -> daher auch keine ermittlung der RSIDs
