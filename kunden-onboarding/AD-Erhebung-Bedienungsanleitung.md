# AD-Erhebung: Bedienungsanleitung

Skript: `AD-Erhebung.ps1` (eine einzige Datei, Version mit Excel- und Word-Erzeugung)

Das Skript liest eine Active-Directory-Domäne **nur lesend** aus und erzeugt daraus in einem neuen Ausgabeordner:

| Datei | Inhalt |
|---|---|
| `Kunden-Onboarding-AD-<Domäne>.xlsx` | Onboarding-Arbeitsmappe (Vorlage mit den Tabs 1 bis 11, leer) plus neun AD-Tabs |
| `AD-Zusammenfassung-<Domäne>.docx` | Word-Zusammenfassung für das Kundengespräch |
| `domaene.csv`, `domain-controller.csv`, `standorte-ad-sites.csv`, `subnetze.csv`, `organisationseinheiten.csv`, `nutzer.csv`, `computerkonten.csv`, `gruppen.csv`, `privilegierte-gruppen-mitglieder.csv` | Rohdaten als CSV (Semikolon-getrennt, UTF-8) |
| `AD-Uebersicht.md`, `AD-Uebersicht.html` | Übersicht in Markdown und als durchsuchbare Webseite |
| `fehlerprotokoll.txt` | Nur wenn Abschnitte nicht gelesen werden konnten |

Die Originaldateien im Kundenordner werden nie verändert. Die Excel-Vorlage ist im Skript eingebettet. Excel und Word werden direkt als Dateien geschrieben: Office, openpyxl oder andere Module sind **nicht** nötig.

---

## 1. Voraussetzungen

### 1.1 Rechner
- Windows 10 oder 11, oder Windows Server mit Zugriff auf die Domäne.
- Windows PowerShell 5.1 (ab Werk vorhanden) oder PowerShell 7 unter Windows.

### 1.2 ActiveDirectory-Modul (RSAT)
Das Skript braucht das Modul `ActiveDirectory`. Ohne Modul läuft es weiter, alle AD-Abschnitte werden aber übersprungen und im Fehlerprotokoll vermerkt.

Installation auf Windows 10/11 (PowerShell **als Administrator**):

```powershell
Add-WindowsCapability -Online -Name "Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0"
```

Auf einem Domain Controller oder Server:

```powershell
Install-WindowsFeature RSAT-AD-PowerShell
```

Prüfen, ob das Modul vorhanden ist:

```powershell
Get-Module -ListAvailable ActiveDirectory
```

### 1.3 Netzwerk
Der Rechner muss einen Domain Controller erreichen: LDAP (389), ADWS (9389) und Kerberos (88).

### 1.4 Berechtigungen
- Das Konto braucht **Leserechte** im AD.
- Für vollständige Gruppen- und Mitgliederdaten (insbesondere der privilegierten Gruppen) ist ein **Domain-Admin-Konto** nötig, das der Kunde für die Erhebung bereitstellt. Ohne diese Rechte fehlen einzelne Gruppen, der Rest der Erhebung läuft weiter.
- Das Skript fragt die Anmeldedaten beim Start per Dialog ab (`Get-Credential`). Passwörter werden nicht im Skript gespeichert.

### 1.5 Keine Änderungen am Netzwerk
Das Skript verwendet ausschließlich `Get-`-Befehle. Es schreibt nichts ins Verzeichnis und ändert keine Konten, Gruppen oder Computer. Auf dem Rechner entsteht nur der Ausgabeordner.

---

## 2. Aufruf

Die Datei in einen Ordner kopieren, PowerShell in diesem Ordner öffnen und ausführen.

### 2.1 Standardaufruf
```powershell
.\AD-Erhebung.ps1 -Domain firma.local
```
Es öffnet sich der Anmeldedialog. Danach entsteht der Ordner `AD-Erhebung-firma.local-<Zeitstempel>` im aktuellen Verzeichnis.

### 2.2 Bestimmter Domain Controller
```powershell
.\AD-Erhebung.ps1 -Domain firma.local -Server dc01.firma.local
```

### 2.3 Eigener Ausgabeordner
```powershell
.\AD-Erhebung.ps1 -Domain firma.local -OutputDir D:\Kunden\Firma\AD
```
Empfehlung: einen neuen, leeren Ordner angeben. Liegen in dem Ordner schon CSV-Dateien mit denselben Namen, werden diese beim Lauf ersetzt.

### 2.4 Eigene Excel-Vorlage
```powershell
.\AD-Erhebung.ps1 -Domain firma.local -TemplatePath C:\Vorlagen\Kunden-Onboarding-Excel.xlsx
```
Ohne diesen Parameter wird die eingebettete Vorlage verwendet. Die angegebene Datei muss eine `.xlsx`-Datei mit den Tabs der Onboarding-Vorlage sein.

### 2.5 Inaktivitätsschwelle ändern
```powershell
.\AD-Erhebung.ps1 -Domain firma.local -InactiveDays 180
```

### 2.6 Alles zusammen
```powershell
.\AD-Erhebung.ps1 -Domain firma.local -Server dc01.firma.local -OutputDir D:\Kunden\Firma\AD -InactiveDays 90
```

### 2.7 Ausführungsrichtlinie
Falls PowerShell das Skript wegen der Ausführungsrichtlinie blockiert, für die aktuelle Sitzung erlauben:
```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```
Das gilt nur für dieses PowerShell-Fenster und ändert keine Systemeinstellung.

---

## 3. Parameter

| Parameter | Pflicht | Standard | Bedeutung |
|---|---|---|---|
| `-Domain` | ja | – | DNS-Name der Domäne, z. B. `firma.local` |
| `-Server` | nein | die Domäne selbst | Bestimmter Domain Controller für alle Abfragen |
| `-OutputDir` | nein | `AD-Erhebung-<Domäne>-<JJJJMMTT-HHmmss>` im aktuellen Verzeichnis | Zielordner für alle Dateien |
| `-InactiveDays` | nein | `90` | Ab dieser Zahl von Tagen ohne Anmeldung (oder nie angemeldet) gilt ein Konto bzw. Computer als inaktiv. Zulässig: 1 bis 3650 |
| `-TemplatePath` | nein | eingebettete Vorlage | Eigene Excel-Vorlage (`.xlsx`) |

---

## 4. Ablauf: was jeder Schritt tut

Das Skript arbeitet die Schritte nacheinander ab. Schlägt ein Schritt fehl, läuft das Skript mit dem nächsten weiter. Der Fehler landet im Fehlerprotokoll.

| Schritt | Was passiert | Ergebnis |
|---|---|---|
| 0. Vorbereitung | Lädt das ActiveDirectory-Modul. Fehlt es, werden alle AD-Schritte übersprungen. | – |
| 0. Anmeldung | Fragt Benutzername und Passwort ab. Ohne Eingabe bricht das Skript ab, ohne etwas abzufragen. | – |
| 1. Domäne und Domain Controller | Liest Domänen- und Gesamtstrukturdaten, Betriebsmaster (Schema, Domänennamen, PDC, RID, Infrastruktur) und alle DCs mit Standort, Betriebssystem, Globalem Katalog und RODC-Status. | `domaene.csv`, `domain-controller.csv` |
| 2. AD-Standorte und Subnetze | Liest die AD-Sites und Subnetze mit ihrer Zuordnung zum Standort. | `standorte-ad-sites.csv`, `subnetze.csv` |
| 3. Organisationseinheiten | Liest alle OUs mit Pfad und Beschreibung. | `organisationseinheiten.csv` |
| 4. Nutzerkonten | Liest Name, Anmeldename, UPN, E-Mail, Abteilung, Funktion, Standort, Status, letzte Anmeldung, Erstellungsdatum und OU. Berechnet „Inaktiv“ anhand von `-InactiveDays`. Passwort- oder Sicherheitsattribute werden nicht angefordert. | `nutzer.csv` |
| 5. Computerkonten | Liest Name, DNS-Name, Betriebssystem, Version, letzte Anmeldung, IPv4, Status und OU. Leitet „Server“ oder „Client“ aus dem Betriebssystem ab. | `computerkonten.csv` |
| 6. Gruppen | Liest alle Gruppen mit Bereich, Typ, Anzahl direkter Mitglieder und OU. | `gruppen.csv` |
| 7. Privilegierte Gruppen | Sucht über feste SIDs: Domänen-Admins, Schema-Admins, Administratoren, Kontooperatoren, Sicherungsoperatoren und Organisations-Admins (Stammdomäne). Listet nur die **direkten** Mitglieder mit Kontonamen und Objektart. Verschachtelte Gruppen werden nicht aufgelöst. | `privilegierte-gruppen-mitglieder.csv` |
| 8. Markdown-Übersicht | Schreibt `AD-Uebersicht.md` mit den Abschnitten zu Domäne, DCs, Standorten, Nutzern, OUs, Administrativen Gruppen, Servern und Endgeräten. | `AD-Uebersicht.md` |
| 9. HTML-Übersicht | Schreibt `AD-Uebersicht.html`: Kennzahlen-Kacheln, Tabellen, Suchfeld, inaktive Konten gelb markiert. Ohne externe Dateien, Internetzugriff oder Skripte. | `AD-Uebersicht.html` |
| 10. Excel-Datei | Kopiert die Vorlage und ergänzt neun AD-Tabs (siehe Abschnitt 5.1). Die Tabs 1 bis 11 bleiben leer. Der Eintrag im Übersichtsblatt wird mit Datum und Domäne ergänzt. | `Kunden-Onboarding-AD-<Domäne>.xlsx` |
| 11. Word-Datei | Erzeugt die Zusammenfassung (siehe Abschnitt 5.3). | `AD-Zusammenfassung-<Domäne>.docx` |
| 12. Fehlerprotokoll | Nur wenn Fehler auftraten: schreibt alle Meldungen in eine Textdatei. | `fehlerprotokoll.txt` |

Das Skript liest nur die Daten. Die Excel- und Word-Dateien entstehen aus diesen Daten im Ausgabeordner, nicht aus dem Netzwerk.

---

## 5. Ausgabe

### 5.1 Excel-Datei `Kunden-Onboarding-AD-<Domäne>.xlsx`

Die Mappe enthält alle Tabs der Vorlage in ihrer Reihenfolge (Übersicht, 1 Ansprechpartner bis 11 Abschluss) plus neun AD-Tabs am Ende:

| Tab | Inhalt |
|---|---|
| AD Kennzahlen | Zusammenfassung mit Anzahlen (siehe Abschnitt 5.4) |
| AD Domäne | Domänen- und Gesamtstrukturdaten, Betriebsmaster |
| AD Domain Controller | Alle DCs |
| AD Standorte | AD-Sites mit Beschreibung |
| AD Subnetze | Subnetze mit AD-Standort und Ort |
| AD Nutzer | Alle Nutzerkonten |
| AD Computer | Alle Computerkonten |
| AD Gruppen | Alle Gruppen |
| AD Privilegierte | Direkte Mitglieder der privilegierten Gruppen |

Jeder AD-Tab hat in Zeile 1 die Überschrift, in Zeile 2 den Stand der Erhebung und ab Zeile 4 die Tabelle mit Kopfzeile. Die Kopfzeile ist fixiert. Leere Zellen bedeuten: im AD nicht gepflegt.

Beispiel (gekürzt, Testdomäne `corp.test`), Tab „AD Nutzer“:

```
AD Nutzerkonten
Stand: 12.11.2026 20:43 (UTC). Leere Felder = im AD nicht gepflegt. Keine Passwörter, Hashes oder Schlüssel.
Name         | Anmeldename | UPN              | E-Mail         | Abteilung   | Funktion | Standort | Status      | Letzte Anmeldung | Inaktiv | Erstellt   | OU
Alt Konto    | altkonto    | alt@corp.test    |                | Buchhaltung |          | Hamburg  | deaktiviert | 2025-10-08 21:42 | ja      | 2015-01-01 | corp.test/Hamburg
Erika Test   | etest       |                  |                |             |          |          | aktiv       |                  | ja      | 2021-05-01 | corp.test/Berlin
Max Muster   | mmuster     | mmuster@corp.test| max@corp.test  | Vertrieb    | Leitung  | Berlin   | aktiv       | 2026-11-10 21:42 | nein    | 2019-03-01 | corp.test/Berlin/Vertrieb
```

Die Zeile „Erika Test“ zeigt ein Konto, das aktiv ist, aber nie angemeldet wurde. Es ist deshalb als inaktiv markiert, und Abteilung und Standort sind leer, weil sie im AD nicht gepflegt sind.

Beispiel, Tab „AD Standorte“:
```
Standort (AD-Site) | Beschreibung
Berlin             |
Hamburg            | Niederlassung Nord
```

### 5.2 CSV-Dateien

Semikolon-getrennt, UTF-8, Kopfzeile in der ersten Zeile. In Excel öffnen: „Daten > Aus Text/CSV“ wählen, Trennzeichen „Semikolon“.

`nutzer.csv` (gekürzt):
```
"Name";"Anmeldename";"UPN";"E-Mail";"Abteilung";"Funktion";"Standort";"Status";"Letzte Anmeldung";"Inaktiv";"Erstellt";"OU"
"Alt Konto";"altkonto";"alt@corp.test";"";"Buchhaltung";"";"Hamburg";"deaktiviert";"10/8/2025 9:43:42 PM";"ja";"1/1/2015 12:00:00 AM";"corp.test/Hamburg"
"Max Muster";"mmuster";"mmuster@corp.test";"max@corp.test";"Vertrieb";"Leitung";"Berlin";"aktiv";"11/10/2026 9:43:42 PM";"nein";"3/1/2019 12:00:00 AM";"corp.test/Berlin/Vertrieb"
```

Hinweis: Datumsangaben in den CSV-Dateien stehen im Format, das PowerShell auf dem Rechner verwendet (hier US-Format). In der Excel-Datei und im Word sind sie einheitlich als `JJJJ-MM-TT` geschrieben.

`privilegierte-gruppen-mitglieder.csv` (gekürzt):
```
"Gruppe";"Mitglied";"Objektart"
"Gruppe S-1-5-21-1111-2222-3333-512";"mmuster";"user"
"Gruppe S-1-5-21-1111-2222-3333-512";"Ops-Team";"group"
```

Gruppen werden in dieser Datei nach dem Namen angezeigt. Ist der Name nicht lesbar, steht die SID dort (wie im Beispiel bei Domänen-Admins und den eingebauten Gruppen).

### 5.3 Word-Datei `AD-Zusammenfassung-<Domäne>.docx`

Gliederung, wie im Dokument zu sehen:

1. **Ansprechpartner**: Tabelle mit den Rollen aus dem Fragebogen (Hauptansprechpartner, Stellvertretung, Geschäftsführung, Notfallkontakt, Externer Dienstleister). Name, E-Mail und Telefon bleiben „offen“, weil das AD diese Angaben nicht enthält. Darunter die direkten Mitglieder der privilegierten Gruppen als Gesprächsgrundlage.
2. **Standorte**: AD-Sites mit Anzahl der Subnetze und DCs, dann die Subnetze und die Nutzer je Standort (AD-Feld „Standort“).
3. **Nutzer**: Anzahl gesamt, aktiv, deaktiviert, inaktiv; Tabelle nach Abteilung.
4. **Systeme und Domäne**: Domänendaten, DCs, Server, Clients nach Betriebssystem, Gruppen.
5. **Offene Punkte**: Fragen, die nur der Kunde beantworten kann, und Hinweise zu fehlenden AD-Feldern.
6. **AD-Kennzahlen**: Tabelle mit den Kennzahlen aus Abschnitt 5.4.

Beispiel (gekürzt, Testdaten), Abschnitt „2. Standorte“:
```
AD-Standorte (Sites): 2. Die Zuordnung zu den Fragebogen-Standorten (Adresse, Ansprechpartner vor Ort, Internetanschluss) ist manuell zu prüfen; Adressen sind im AD nicht hinterlegt.

AD-Standort | Beschreibung        | Subnetze | Domain Controller
Berlin      |                     | 1        | 1
Hamburg     | Niederlassung Nord  | 1        | 1
```

### 5.4 Kennzahlen (Tab „AD Kennzahlen“ und Word-Abschnitt 6)

Beispiel (Testdaten):
```
Kennzahl                                              | Wert
Domäne (DNS-Name)                                     | corp.test
Domain Controller                                     | 2
  davon Globaler Katalog                              | 2
  davon schreibgeschützt (RODC)                       | 1
AD-Standorte (Sites)                                  | 2
Subnetze                                              | 2
Nutzerkonten gesamt                                   | 3
  aktiv                                               | 2
  deaktiviert                                         | 1
  aktiv, aber inaktiv (über 90 Tage oder nie angemeldet) | 1
Computerkonten gesamt                                 | 3
  davon Server                                        | 1
  davon Clients                                       | 2
Gruppen gesamt                                        | 3
Direkte Mitglieder in privilegierten Gruppen          | 10
Erhebungsfehler / übersprungene Abschnitte            | 1
```

Wird ein Abschnitt nicht gelesen, steht bei seinen Kennzahlen „nicht erhoben“ statt einer Zahl.

### 5.5 Markdown- und HTML-Übersicht

`AD-Uebersicht.md` gliedert sich nach dem Fragebogen (Abschnitte 2, 3, 4.2, 5.2, 8). Ausschnitt aus dem Testlauf:

```
# AD-Erhebung: corp.test

## Domäne (Fragebogen 4.2)

| Eigenschaft | Wert |
|---|---|
| DNS-Name | corp.test |
| NetBIOS-Name | CORP |
| Domänenfunktionsebene | Windows2016Domain |

## Domain Controller (Fragebogen 5 und 6.1)

| Name | Hostname | IPv4 | AD-Standort | Betriebssystem | Globaler Katalog | Schreibgeschützt (RODC) |
|---|---|---|---|---|---|---|
| DC01 | dc01.corp.test | 10.0.0.10 | Berlin | Windows Server 2022 | ja | nein |
| DC02 | dc02.corp.test | 10.1.0.10 | Hamburg | Windows Server 2019 | ja | ja |
```

`AD-Uebersicht.html` mit einem Doppelklick im Browser öffnen. Das Suchfeld oben filtert alle Tabellen. Inaktive Konten sind gelb markiert.

### 5.6 Fehlerprotokoll `fehlerprotokoll.txt`
Liegt nur vor, wenn etwas nicht gelesen werden konnte. Beispiel (Testlauf mit absichtlich verweigerter Gruppe):
```
[Privilegierte Gruppen] Gruppe S-1-5-21-1111-2222-3333-518 nicht gelesen: Zugriff verweigert (Test)
```
Dieses Protokoll steht auch in der Übersicht und in der Word-Datei unter „Offene Punkte“.

---

## 6. Typische Fehler und Lösungen

| Symptom | Ursache | Lösung |
|---|---|---|
| `Das Modul ActiveDirectory ist nicht verfügbar` / alle AD-Abschnitte im Fehlerprotokoll als „übersprungen“ | RSAT fehlt | Installationsbefehl aus 1.2 als Administrator ausführen, neues PowerShell-Fenster öffnen |
| `Abbruch: Es wurden keine Anmeldedaten angegeben` | Anmeldedialog geschlossen | Skript erneut starten und Zugangsdaten eingeben |
| `Der Server ist nicht erreichbar` oder Zeitüberschreitung in jedem Abschnitt | Kein Domain Controller erreichbar | Netzwerk und DNS prüfen, mit `-Server` einen bekannten DC angeben, Ports 389, 9389 und 88 freigeben |
| `Zugriff verweigert` bei einer oder mehreren privilegierten Gruppen | Konto hat zu wenig Rechte | Domain-Admin-Konto des Kunden verwenden (Abschnitt 1.4) |
| Einzelne Gruppen fehlen in der Liste der privilegierten Gruppen | Gruppe existiert nicht (z. B. Organisations-Admins in einer Einzeldomäne) oder Lesezugriff fehlt | Fehlerprotokoll prüfen; fehlt eine Gruppe, ist das bei Einzeldomänen normal |
| Excel- oder Word-Datei im Ausgabeordner mit Zusatz `-HHmmss` im Namen | Im Ausgabeordner gab es schon eine Datei mit diesem Namen. Das Skript überschreibt nie und hängt deshalb die Uhrzeit an | Nichts tun; die neue Datei ist die aktuelle |
| Umlaute werden in der CSV falsch angezeigt | Die CSV wird mit einer anderen Kodierung geöffnet | Mit „Daten > Aus Text/CSV“ öffnen und Kodierung „UTF-8“ wählen |
| `Die Ausführung von Skripts ist auf diesem System deaktiviert` | Ausführungsrichtlinie | Befehl aus 2.7 verwenden |
| Lauf dauert lange | Die Domäne hat viele Konten und Gruppen | Das Skript arbeitet ohne Eingabe weiter; auf die Meldung „Fertig“ warten |

Wenn ein Schritt fehlschlägt, bleibt der Lauf erhalten: alle anderen Dateien sind trotzdem da. Vor dem Weitergeben immer `fehlerprotokoll.txt` prüfen.

---

## 7. Bewusst nicht erhoben

Das Skript erhebt **keine** der folgenden Daten:

- Passwörter, Passworthashes oder Kerberos-Schlüssel
- LAPS-Passwörter und LAPS-Schlüssel
- BitLocker-Wiederherstellungsschlüssel
- Gruppenmitglieder aus verschachtelten Gruppen (es werden nur direkte Mitglieder gelesen)
- Inhalte von Dateien, Freigaben oder Postfächern
- Anmeldeverlauf oder Ereignisprotokolle

Alle Abfragen sind lesend (`Get-`). Es werden keine Konten, Gruppen, Computer, Richtlinien oder DNS-Einträge verändert, und es wird nichts an Dritte gesendet.

---

## 8. Nach dem Lauf: Prüfliste

1. Fehlerprotokoll prüfen, falls vorhanden.
2. Excel-Datei öffnen: Tab „AD Kennzahlen“ auf Plausibilität prüfen (Anzahl Nutzer und Computer stimmt mit dem Kunden überein?).
3. Standorte: Die AD-Sites mit den Fragebogen-Standorten (Abschnitt 2) abgleichen. Die Zuordnung macht das Skript nicht automatisch.
4. Felder „Abteilung“ und „Standort“ im AD sind oft nicht gepflegt. Lücken mit dem Kunden klären.
5. Inaktive Konten (z. B. deaktiviert oder seit mehr als 90 Tagen nicht angemeldet) mit dem Kunden besprechen.
6. Word-Datei als Grundlage für das Gespräch verwenden und die Kundendaten in Excel-Tabs 1 bis 11 eintragen.

---

## 9. Ehrlicher Teststand

- **Syntax**: Das Skript wurde mit PowerShell 7.6 (pwsh) geparst, Ergebnis: SYNTAX OK.
- **Excel- und Word-Erzeugung**: Mit Testdaten (Domäne `corp.test`, Testmodul statt echtem AD) ausgeführt. Die Excel-Datei wurde mit openpyxl geöffnet und alle AD-Tabs gelesen, die Word-Datei mit python-docx (38 Absätze, 11 Tabellen). Beide Dateien sind als XML wohlgeformt. Die eingebettete Vorlage ist byte-identisch mit der Originaldatei.
- **Nicht getestet**: Ein echter Lauf gegen eine Domäne war nicht möglich, weil von dem Testrechner aus kein Domain Controller erreichbar ist. Ebenso nicht getestet: Windows PowerShell 5.1 und Microsoft Excel/Word selbst. Beim ersten Lauf beim Kunden die CSV-Dateien und die Excel-Datei kurz gegenprüfen.
- **Vorlagen**: Die Excel-Vorlage des Fragebogens hat die Tabs 1 bis 11 unverändert behalten.
