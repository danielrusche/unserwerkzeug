# IT-Onboarding: Fragebogen und Checkliste

**Kunde:** ____________________________________
**Ansprechpartner (Kunde):** ____________________________________
**Onboarding durchgeführt von:** ____________________________________
**Datum:** ____________________________________

---

## Inhaltsverzeichnis

1. Kundenansprechpartner
2. Standorte
3. Nutzer
4. Administrative Zugänge
5. Systeme
6. Netzwerk
7. Internetanbindung
8. Endgeräte
9. Hersteller und Produkte
10. Bestandsdokumentation
11. Abschluss (Checkliste, offene Punkte, Unterschriften)

---

## Hinweise zum Ausfüllen

- Bitte **keine Passwörter** oder andere Geheimnisse in dieses Dokument eintragen. Bei Zugangsdaten wird nur vermerkt, **wo** das Passwort hinterlegt ist (z. B. Passwortmanager, Tresor, versiegelter Umschlag).
- Wenn eine Angabe nicht bekannt ist, bitte „unbekannt“ eintragen statt leer zu lassen.
- Die Abschnitte sind so aufgebaut, dass jeder für sich ausgefüllt werden kann. Standorte werden nur in Abschnitt 2 erfasst; in allen anderen Abschnitten wird auf diese Namen verwiesen. Ebenso werden Server nur in Abschnitt 5 erfasst und in Abschnitt 8 nur als Verweis genannt.
- Bitte jeweils Name, Datum und Kürzel der ausfüllenden Person angeben.

---

## 1. Kundenansprechpartner

Ziel: Ansprechpartner für Rückfragen, Störungen und Freigaben.

| Rolle | Name | E-Mail | Telefon (Büro) | Telefon (mobil) | Adresse | Verfügbarkeit |
|---|---|---|---|---|---|---|
| Hauptansprechpartner | | | | | | |
| Stellvertretung | | | | | | |
| Geschäftsführung / Entscheider | | | | | | |
| Ansprechpartner Notfall (außerhalb der Geschäftszeiten) | | | | | | |
| Externer Dienstleister (z. B. Telefonie, Internet, Drucker) | | | | | | |

Bei der Adresse der Ansprechpartner bitte den Hauptsitz angeben; weitere Standorte stehen in Abschnitt 2.

**Fragen:**
- Wer darf Bestellungen und Änderungen freigeben?
- Wer darf Störungsmeldungen im Namen des Kunden aufgeben?
- Welche Kommunikationswege sind gewünscht (Telefon, E-Mail, Ticketsystem)?

---

## 2. Standorte

Ziel: Übersicht aller Standorte mit Ansprechpartner, Nutzerzahl und dort vorhandenen Systemen. Die Standortnamen aus dieser Tabelle werden in den Abschnitten 3, 5, 6, 7 und 8 verwendet.

| Standort | Adresse | Ansprechpartner vor Ort | Anzahl Nutzer | Systeme am Standort (Details in Abschnitt 5) | Internetanschluss vorhanden (ja / nein) |
|---|---|---|---|---|---|
| | | | | | |
| | | | | | |

**Fragen:**
- Wie viele Standorte gibt es? Gibt es Außenstellen, Home-Office-Arbeitsplätze oder mobile Mitarbeiter?
- Gibt es einen Serverraum oder Technikraum je Standort? Wo steht er, wer hat Zutritt?
- Welche Standorte sind über VPN oder Standleitung verbunden? (Details in Abschnitt 7.3)

---

## 3. Nutzer

Ziel: Erfassung aller Benutzer mit Standort und Abteilung. Der Standort wird aus Abschnitt 2 übernommen. Bitte bei größeren Benutzerzahlen eine Exportdatei aus dem Verzeichnisdienst beilegen (Name, Abteilung, Standort; keine Passwörter).

| Name | Benutzername / Anmeldename | E-Mail-Adresse | Standort (aus Abschnitt 2) | Abteilung | Rolle / Funktion | Berechtigungsgruppe |
|---|---|---|---|---|---|---|
| | | | | | | |
| | | | | | | |

**Fragen:**
- Wie viele Nutzer gibt es insgesamt?
- Welche Abteilungen gibt es, und wer leitet sie?
- Gibt es Gäste, Praktikanten oder externe Dienstleister mit Zugang? Wie lange ist deren Zugang gültig?
- Wie werden neue Nutzer angelegt und Austritte gemeldet? Wer macht das?

---

## 4. Administrative Zugänge

Ziel: Überblick über alle privilegierten Zugänge. Es werden nur Benutzernamen und der Ablageort des Passworts erfasst.

### 4.1 Lokale Administratoren

| Gerät / System | Benutzername | Ablageort des Passworts | Bemerkung |
|---|---|---|---|
| | | | |
| | | | |
| | | | |

**Fragen:**
- Gibt es auf den Clients und Servern eigene lokale Admin-Konten? Auf welchen Geräten?
- Wurden Standardkonten (z. B. „Administrator“, „root“) umbenannt oder deaktiviert?

### 4.2 Domain Admin und Active Directory

| Konto | Benutzername | Ablageort des Passworts | Bemerkung |
|---|---|---|---|
| Domain Admin | | | |
| Weiteres Admin-Konto | | | |
| Dienstkonto(e) | | | |

**Fragen:**
- Wie heißt die Domäne (intern und ggf. öffentlich)?
- Welche Konten sind Mitglied in den Gruppen „Domain Admins“ und „Enterprise Admins“?
- Gibt es einen Dienstkonto-Bestand, dessen Passwörter regelmäßig gewechselt werden müssen?

### 4.3 Management-Zugänge (iDRAC, iLO, IPMI, KVM, Web-Oberflächen)

| System / Gerät | Art (iDRAC, iLO, IPMI, Web-UI, ...) | Adresse oder Hostname | Benutzername | Ablageort des Passworts |
|---|---|---|---|---|
| | | | | |
| | | | | |
| | | | | |

Die Systeme stammen aus Abschnitt 5.

**Fragen:**
- Welche Server haben eine Management-Karte (iDRAC, iLO, IPMI)?
- Sind die Management-Schnittstellen über ein eigenes Management-Netz erreichbar?
- Gibt es Zugänge zu Hypervisor-Oberflächen (z. B. Proxmox, vCenter, Hyper-V-Manager)?

### 4.4 Cloud- und Dienstkonten

| Dienst | Benutzername / Konto | Ablageort des Passworts | Zwei-Faktor-Methode |
|---|---|---|---|
| Microsoft 365 / Google Workspace | | | |
| Lizenzierung / Hersteller-Portale | | | |
| Backup-Dienst | | | |
| Sonstige | | | |

---

## 5. Systeme

Ziel: Vollständige Liste aller Systeme. Pro System werden Hersteller, Produktname, Standort (aus Abschnitt 2), Rack und Höheneinheit (HE) erfasst. Falls kein Rack vorhanden ist, bitte „kein Rack“ eintragen.

### 5.1 Virtualisierung und Hypervisoren

| Bezeichnung / Hostname | Hersteller | Produktname | Standort | Rack | HE | Bemerkung |
|---|---|---|---|---|---|---|
| | | | | | | |
| | | | | | | |

**Fragen:**
- Welche Virtualisierungsplattform wird eingesetzt (z. B. Proxmox VE, VMware ESXi, Hyper-V)?
- Wie viele Hosts (Nodes) gibt es, und sind sie zu einem Cluster verbunden?
- Welche wichtigen virtuellen Maschinen laufen dort (Name, Zweck)?

### 5.2 Physische Server

| Hostname | Hersteller | Produktname / Modell | Standort | Rack | HE | Betriebssystem | Bemerkung |
|---|---|---|---|---|---|---|---|
| | | | | | | | |
| | | | | | | | |

**Fragen:**
- Welche Rollen haben die Server (Domain Controller, Dateiserver, Anwendungsserver, Datenbank, ...)?
- Welches Betriebssystem und welche Version läuft darauf?

### 5.3 Proxmox-Hosts

| Hostname | Version | Standort | Rack | HE | Cluster-Name | Bemerkung |
|---|---|---|---|---|---|---|
| | | | | | | |

### 5.4 Drucker und Druckserver

| Gerät | Hersteller | Produktname | Standort | Raum / Abteilung | IP-Adresse | Druckserver (ja / nein) |
|---|---|---|---|---|---|---|
| | | | | | | |
| | | | | | | |

**Fragen:**
- Gibt es einen zentralen Druckserver? Welche Software wird dort genutzt?
- Welche Drucker sind über Treiber der Clients angebunden, welche über das Netzwerk?
- Gibt es Multifunktionsgeräte mit Scan-to-Mail oder Scan-to-Ordner?

---

## 6. Netzwerk

Ziel: DNS, DHCP, IP-Bereiche, VLANs, Namenskonventionen und Switche. Standorte werden aus Abschnitt 2 übernommen. Die Internetanbindung (Provider, Firewall, VPN, Gateway) steht in Abschnitt 7.

### 6.1 DNS

| Zone / Domäne | DNS-Server (primär / sekundär) | Hersteller / Software | Interne Zone (ja / nein) | Öffentliche Zone (ja / nein) | Bemerkung |
|---|---|---|---|---|---|
| | | | | | |

**Fragen:**
- Welche DNS-Server sind im Einsatz (z. B. Windows-DNS, BIND, Pi-hole, Router)?
- Wer verwaltet die öffentlichen DNS-Einträge (Registrar, Provider, Hosting)?
- Gibt es Weiterleitungen (Forwarder) auf bestimmte Server?

### 6.2 DHCP

| Standort / Netz | DHCP-Server | Bereich (von–bis) | Gateway | Lease-Dauer | Reservierungen (ja / nein) |
|---|---|---|---|---|---|
| | | | | | |

**Fragen:**
- Wo läuft der DHCP-Dienst (Firewall, Server, Router)?
- Gibt es feste IP-Adressen (Reservierungen) für Server, Drucker oder Telefone? Wo sind sie dokumentiert?

### 6.3 IP-Adressbereiche und VLANs

| Netz / VLAN | Bezeichnung | Adressbereich (z. B. 10.0.10.0/24) | Standort | Zweck (Server, Clients, Gäste, Management, IoT) | Gateway |
|---|---|---|---|---|---|
| | | | | | |
| | | | | | |

**Fragen:**
- Welche Subnetze sind im Einsatz? Gibt es Überschneidungen zwischen Standorten?
- Gibt es ein separates Gäste-WLAN oder ein Management-Netz?
- Welche Netze sind über VPN oder Site-to-Site-Verbindungen miteinander verbunden? (Details in Abschnitt 7.3)

### 6.4 Namenskonventionen

| Gerätetyp | Vorgegebene Schreibweise / Muster (z. B. SRV-STANDORT-01) | Bemerkung |
|---|---|---|
| Server | | |
| Clients / Notebooks | | |
| Drucker | | |
| Netzwerkgeräte | | |
| Benutzerkonten | | |

**Fragen:**
- Gibt es eine bestehende Namenskonvention? Falls ja, bitte Beispiele nennen.
- Sollen neue Geräte nach einer vorgegebenen Konvention benannt werden?

### 6.5 Switche

| Standort | Hersteller | Modell | Anzahl Ports | Firmware-Version | Rack | HE | Verwaltbar (ja / nein) |
|---|---|---|---|---|---|---|---|
| | | | | | | | |
| | | | | | | | |

**Fragen:**
- Sind die Switche verwaltbar (Managed) und in welcher Konfiguration?
- Gibt es VLAN-Trunks oder Stacking zwischen den Switchen?

---

## 7. Internetanbindung

Ziel: Anschluss, Firewall, VPN und Gateway/Router. Standorte werden aus Abschnitt 2 übernommen, Admin-Zugänge zur Firewall stehen in Abschnitt 4.

### 7.1 Provider und Anschluss

| Standort | Provider | Anschlusstyp (DSL, Glasfaser, Kabel, LTE/5G, Standleitung) | Bandbreite (down / up) | Vertragsnummer / Kundennummer | Feste IP-Adressen (ja / nein) |
|---|---|---|---|---|---|
| | | | | | |

**Fragen:**
- Gibt es einen Backup-Anschluss (Failover)?
- Wann läuft der Vertrag aus, und gibt es eine Kündigungsfrist?

### 7.2 Firewall

| Standort | Hersteller | Modell | Firmware-Version | Lizenz / Support bis | Standort im Rack (Rack, HE) | Admin-Zugang (siehe Abschnitt 4) |
|---|---|---|---|---|---|---|
| | | | | | | |

**Fragen:**
- Welche Regeln sind kritisch (z. B. Port-Weiterleitungen, Zugriff auf Server von außen)?
- Wann wurde die Firmware zuletzt aktualisiert?

### 7.3 VPN

| Standort / Verbindung | Typ (Site-to-Site, Client-VPN) | Hersteller / Software | Benutzer oder Gruppen | Authentifizierung (z. B. Zertifikat, MFA) | Bemerkung |
|---|---|---|---|---|---|
| | | | | | |

**Fragen:**
- Welche Nutzer benötigen VPN-Zugang?
- Wie erfolgt die Anmeldung (Benutzername und Passwort, Zertifikat, Zwei-Faktor)?

### 7.4 Gateway und Router

| Standort | Gerät (Hersteller, Modell) | Firmware-Version | Rack / HE | Bemerkung |
|---|---|---|---|---|
| | | | | |

---

## 8. Endgeräte

Ziel: Zuordnung der Clients und Geräte zu Nutzern (aus Abschnitt 3) und Standorten (aus Abschnitt 2). Server sind in Abschnitt 5 erfasst und werden hier nur referenziert; Netzwerkgeräte stehen in Abschnitt 6.5.

| Gerätename | Typ (Desktop, Notebook, Tablet, Smartphone, Server-Verweis) | Hersteller | Modell | Betriebssystem | Zugewiesener Nutzer | Standort | Seriennummer (optional) |
|---|---|---|---|---|---|---|---|
| | | | | | | | |
| | | | | | | | |

**Fragen:**
- Wie viele Clients gibt es insgesamt?
- Sind Geräte bereits in einem Verzeichnisdienst oder einer Verwaltungssoftware (MDM, RMM) erfasst?
- Gibt es private Geräte (BYOD), die Zugriff auf Firmendaten haben?
- Welche Geräte sind geleast oder haben einen Wartungsvertrag, und bis wann?

---

## 9. Hersteller und Produkte

Ziel: Übersicht aller im Einsatz befindlichen Hersteller und Produkte. Diese Liste dient als Grundlage für Lizenz-, Support- und Wartungsplanung. Die Hardware-Details stehen in den Abschnitten 5 bis 8.

| Hersteller | Produktname | Kategorie (Server, Netzwerk, Software, Cloud, Sicherheit, Drucker, Sonstiges) | Version | Lizenz / Support bis | Bemerkung |
|---|---|---|---|---|---|
| | | | | | |
| | | | | | |
| | | | | | |

**Fragen:**
- Welche Betriebssysteme sind im Einsatz (Server und Clients)?
- Welche Fachanwendungen werden genutzt, und wer betreut sie beim Hersteller?
- Welche Software-Lizenzen laufen bald ab?

---

## 10. Bestandsdokumentation

Ziel: Feststellen, welche bestehende Dokumentation vorhanden ist und wo sie liegt. Die Dokumente selbst werden nicht in dieses Formular kopiert.

| Dokument | Art (Netzplan, Bestandsliste, Verkabelungsplan, Lizenzübersicht, Notfallhandbuch, Sonstiges) | Vorhanden (ja / nein / teilweise) | Ablageort (Pfad, Ordner, Link, Schrank) | Aktualität (Stand) | Bemerkung |
|---|---|---|---|---|---|
| | | | | | |
| | | | | | |
| | | | | | |

**Fragen:**
- Gibt es einen aktuellen Netzplan und eine Bestandsliste?
- Wo liegen Verkabelungspläne, Rack-Pläne und Dokumentation der Serverkonfiguration?
- Gibt es ein Notfall- oder Wiederherstellungshandbuch, und wann wurde es zuletzt getestet?
- Gibt es ein Backup-Konzept, und wo ist es dokumentiert?

---

## 11. Abschluss

### 11.1 Checkliste

- [ ] Alle Abschnitte 1 bis 10 sind bearbeitet oder mit „unbekannt“ markiert.
- [ ] Keine Passwörter oder Geheimnisse in diesem Dokument eingetragen.
- [ ] Offene Punkte sind notiert und einem Verantwortlichen zugeordnet.
- [ ] Fragebogen mit dem Kunden besprochen.

### 11.2 Offene Punkte

| Nr. | Thema | Verantwortlich | Fällig bis |
|---|---|---|---|
| | | | |
| | | | |

### 11.3 Unterschriften

**Unterschrift Kunde:** ____________________________________ **Datum:** ____________

**Unterschrift Onboarding-Mitarbeiter:** ____________________________________ **Datum:** ____________
