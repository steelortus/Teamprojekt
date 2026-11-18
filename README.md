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

## Thinking for later:
- [ ] Was ist versionsnummer? Intern öffentlich identisch? Wann auf den markt gekommen
- [x] Datastructure for changes across versions and versions across time -> multiple CSVs
- [x] change metadata to test calculations vs read (number of words)
- [ ] welchen Zeiträumen wurden welche Sprachkennzeichnungen




## Issues:
- [ ] No Idea, how things are being counted. Absätze sollten <w:p> sein aber sind nciht.
    - `test_long.docx` 133/35
    - `ProccessorAgreement.docx` 579/58
    - `SoftwareOrder.docx` 43/1
- [ ] RSID sind zu unterschiedlich `ProccessorAgreement.docx`: 170 settings.xml vs 48 document parts -> daher auch keine ermittlung der RSIDs

## Data files:
- `scripts/compareOutput/officeReleases.csv` - Office version -> release date.
- `scripts/compareOutput/languageTags.csv` - language tag -> first Windows version/date (MS-LCID Appendix A <8>), availability in the Office 2013/2016 language tables, and deprecated/legacy tags with their preferred successor (IANA subtag registry, CLDR aliases). `FirstWindowsDate` is the approximate general-availability date of that Windows release. An empty `Office2013`/`Office2016` cell means "not confirmed in that version's localized-language list", not "unavailable".
