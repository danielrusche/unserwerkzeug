# IT-Onboarding: Fragebogen und Checkliste

**Kunde:** ____________________________________
**Ansprechpartner (Kunde):** ____________________________________
**Onboarding durchgeführt von:** ____________________________________
**Datum:** ____________________________________

## Hinweise zum Ausfüllen

- Bitte **keine Passwörter** oder andere Geheimnisse in dieses Dokument eintragen. Bei den Zugangsdaten wird nur vermerkt, **wo** das Passwort hinterlegt ist (z. B. Passwortmanager, Tresor, versiegelter Umschlag).
- Wenn eine Angabe nicht bekannt ist, bitte „unbekannt“ eintragen statt leer zu lassen.
- Die Reihenfolge der Abschnitte ist bewusst gewählt: Zuerst Zugänge und Systeme, dann Standorte, Nutzer und Netzwerk, danach Endgeräte, Internet und Hersteller, zuletzt Ansprechpartner und Dokumentation.
- Bitte jeweils Name, Datum und Kürzel der ausfüllenden Person angeben.

---

## Abschnitt 1: Administrative Zugänge

Ziel: Überblick über alle privilegierten Zugänge. Es werden nur Benutzernamen und der Ablageort des Passworts erfasst.

### 1.1 Lokale Administratoren

| Gerät / System | Benutzername | Ablageort des Passworts | Bemerkung |
|---|---|---|---|
| | | | |
| | | | |
| | | | |

**Fragen:**
- Gibt es auf den Clients und Servern eigene lokale Admin-Konten? Auf welchen Geräten?
- Wurden Standardkonten (z. B. „Administrator“, „root“) umbenannt oder deaktiviert?

### 1.2 Domain Admin und Active Directory

| Konto | Benutzername | Ablageort des Passworts | Bemerkung |
|---|---|---|---|
| Domain Admin | | | |
| Weiteres Admin-Konto | | | |
| Dienstkonto(e) | | | |

**Fragen:**
- Wie heißt die Domäne (intern und ggf. öffentlich)?
- Welche Konten sind Mitglied in den Gruppen „Domain Admins“ und „Enterprise Admins“?
- Gibt es einen Dienstkonto-Bestand, dessen Passwörter regelmäßig gewechselt werden müssen?

### 1.3 Management-Zugänge (iDRAC, iLO, IPMI, KVM, Web-Oberflächen)

| System / Gerät | Art (iDRAC, iLO, IPMI, Web-UI, ...) | Adresse oder Hostname | Benutzername | Ablageort des Passworts |
|---|---|---|---|---|
| | | | | |
| | | | | |
| | | | | |

**Fragen:**
- Welche Server haben eine Management-Karte (iDRAC, iLO, IPMI)?
- Sind die Management-Schnittstellen über ein eigenes Management-Netz erreichbar?
- Gibt es Zugänge zu Hypervisor-Oberflächen (z. B. Proxmox, vCenter, Hyper-V-Manager)?

### 1.4 Cloud- und Dienstkonten

| Dienst | Benutzername / Konto | Ablageort des Passworts | Zwei-Faktor-Methode |
|---|---|---|---|
| Microsoft 365 / Google Workspace | | | |
| Lizenzierung / Hersteller-Portale | | | |
| Backup-Dienst | | | |
| Sonstige | | | |

---

## Abschnitt 2: Systeme

Ziel: Vollständige Liste aller Systeme. Pro System werden Hersteller, Produktname, Standort, Rack und Höheneinheit (HE) erfasst. Falls kein Rack vorhanden ist, bitte „kein Rack“ eintragen.

### 2.1 Virtualisierung und Hypervisoren

| Bezeichnung / Hostname | Hersteller | Produktname | Standort | Rack | HE | Bemerkung |
|---|---|---|---|---|---|---|
| | | | | | | |
| | | | | | | |

**Fragen:**
- Welche Virtualisierungsplattform wird eingesetzt (z. B. Proxmox VE, VMware ESXi, Hyper-V)?
- Wie viele Hosts (Nodes) gibt es, und sind sie zu einem Cluster verbunden?
- Welche wichtigen virtuellen Maschinen laufen dort (Name, Zweck)?

### 2.2 Physische Server

| Hostname | Hersteller | Produktname / Modell | Standort | Rack | HE | Betriebssystem | Bemerkung |
|---|---|---|---|---|---|---|---|
| | | | | | | | |
| | | | | | | | |

**Fragen:**
- Welche Rollen haben die Server (Domain Controller, Dateiserver, Anwendungsserver, Datenbank, ...)?
- Welches Betriebssystem und welche Version läuft darauf?

### 2.3 Proxmox-Hosts

| Hostname | Version | Standort | Rack | HE | Cluster-Name | Bemerkung |
|---|---|---|---|---|---|---|
| | | | | | | |

### 2.4 Drucker und Druckserver

| Gerät | Hersteller | Produktname | Standort | Raum / Abteilung | IP-Adresse | Druckserver (ja / nein) |
|---|---|---|---|---|---|---|
| | | | | | | |
| | | | | | | |

**Fragen:**
- Gibt es einen zentralen Druckserver? Welche Software wird dort genutzt?
- Welche Drucker sind über Treiber der Clients angebunden, welche über das Netzwerk?
- Gibt es Multifunktionsgeräte mit Scan-to-Mail oder Scan-to-Ordner?

---

## Abschnitt 3: Standorte

Ziel: Übersicht aller Standorte und der dort vorhandenen Systeme und Nutzer.

| Standort | Adresse | Ansprechpartner vor Ort | Anzahl Nutzer | Systeme am Standort | Internetanschluss vorhanden (ja / nein) |
|---|---|---|---|---|---|
| | | | | | |
| | | | | | |

**Fragen:**
- Wie viele Standorte gibt es? Gibt es Außenstellen, Home-Office-Arbeitsplätze oder mobile Mitarbeiter?
- Gibt es einen Serverraum oder Technikraum je Standort? Wo steht er, wer hat Zutritt?
- Welche Standorte sind über VPN oder Standleitung verbunden?

---

## Abschnitt 4: Nutzer

Ziel: Erfassung aller Benutzer mit Standort und Abteilung. Bitte bei größeren Benutzerzahlen eine Exportdatei aus dem Verzeichnisdienst beilegen (Name, Abteilung, Standort; keine Passwörter).

| Name | Benutzername / Anmeldename | E-Mail-Adresse | Standort | Abteilung | Rolle / Funktion | Berechtigungsgruppe |
|---|---|---|---|---|---|---|
| | | | | | | |
| | | | | | | |

**Fragen:**
- Wie viele Nutzer gibt es insgesamt?
- Welche Abteilungen gibt es, und wer leitet sie?
- Gibt es Gäste, Praktikanten oder externe Dienstleister mit Zugang? Wie lange ist deren Zugang gültig?
- Wie werden neue Nutzer angelegt und Austritte gemeldet? Wer macht das?

---

## Abschnitt 5: Netzwerk

### 5.1 DNS

| Zone / Domäne | DNS-Server (primär / sekundär) | Hersteller / Software | Interne Zone (ja / nein) | Öffentliche Zone (ja / nein) | Bemerkung |
|---|---|---|---|---|---|
| | | | | | |

**Fragen:**
- Welche DNS-Server sind im Einsatz (z. B. Windows-DNS, BIND, Pi-hole, Router)?
- Wer verwaltet die öffentlichen DNS-Einträge (Registrar, Provider, Hosting)?
- Gibt es Weiterleitungen (Forwarder) auf bestimmte Server?

### 5.2 DHCP

| Standort / Netz | DHCP-Server | Bereich (von–bis) | Gateway | Lease-Dauer | Reservierungen (ja / nein) |
|---|---|---|---|---|---|
| | | | | | |

**Fragen:**
- Wo läuft der DHCP-Dienst (Firewall, Server, Router)?
- Gibt es feste IP-Adressen (Reservierungen) für Server, Drucker oder Telefone? Wo sind sie dokumentiert?

### 5.3 IP-Adressbereiche und VLANs

| Netz / VLAN | Bezeichnung | Adressbereich (z. B. 10.0.10.0/24) | Standort | Zweck (Server, Clients, Gäste, Management, IoT) | Gateway |
|---|---|---|---|---|---|
| | | | | | |
| | | | | | |

**Fragen:**
- Welche Subnetze sind im Einsatz? Gibt es Überschneidungen zwischen Standorten?
- Gibt es ein separates Gäste-WLAN oder ein Management-Netz?
- Welche Netze sind über VPN oder Site-to-Site-Verbindungen miteinander verbunden?

### 5.4 Namenskonventionen

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

---

## Abschnitt 6: Endgeräte

Ziel: Zuordnung der Clients und Geräte zu Nutzern und Standorten.

| Gerätename | Typ (Desktop, Notebook, Tablet, Smartphone) | Hersteller | Modell | Betriebssystem | Zugewiesener Nutzer | Standort | Seriennummer (optional) |
|---|---|---|---|---|---|---|---|
| | | | | | | | |
| | | | | | | | |

**Fragen:**
- Wie viele Clients gibt es insgesamt?
- Sind Geräte bereits in einem Verzeichnisdienst oder einer Verwaltungssoftware (MDM, RMM) erfasst?
- Gibt es private Geräte (BYOD), die Zugriff auf Firmendaten haben?
- Welche Geräte sind geleast oder haben einen Wartungsvertrag, und bis wann?

---

## Abschnitt 7: Internetanbindung

### 7.1 Provider und Anschluss

| Standort | Provider | Anschlusstyp (DSL, Glasfaser, Kabel, LTE/5G, Standleitung) | Bandbreite (down / up) | Vertragsnummer / Kundennummer | Feste IP-Adressen (ja / nein) |
|---|---|---|---|---|---|
| | | | | | |

**Fragen:**
- Gibt es einen Backup-Anschluss (Failover)?
- Wann läuft der Vertrag aus, und gibt es eine Kündigungsfrist?

### 7.2 Firewall

| Standort | Hersteller | Modell | Firmware-Version | Lizenz / Support bis | Standort im Rack (Rack, HE) | Admin-Zugang (siehe Abschnitt 1) |
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

### 7.5 Switche

| Standort | Hersteller | Modell | Anzahl Ports | Firmware-Version | Rack | HE | Verwaltbar (ja / nein) |
|---|---|---|---|---|---|---|---|
| | | | | | | | |
| | | | | | | | |

**Fragen:**
- Sind die Switche verwaltbar (Managed) und in welcher Konfiguration?
- Gibt es VLAN-Trunks oder Stacking zwischen den Switchen?

---

## Abschnitt 8: Hersteller und Produkte

Ziel: Übersicht aller im Einsatz befindlichen Hersteller und Produkte. Diese Liste dient als Grundlage für Lizenz-, Support- und Wartungsplanung.

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

## Abschnitt 9: Kundenansprechpartner

Ziel: Ansprechpartner für Rückfragen, Störungen und Freigaben.

| Rolle | Name | E-Mail | Telefon (Büro) | Telefon (mobil) | Adresse | Verfügbarkeit |
|---|---|---|---|---|---|---|
| Hauptansprechpartner | | | | | | |
| Stellvertretung | | | | | | |
| Geschäftsführung / Entscheider | | | | | | |
| Ansprechpartner Notfall (außerhalb der Geschäftszeiten) | | | | | | |
| Externer Dienstleister (z. B. Telefonie, Internet, Drucker) | | | | | | |

**Fragen:**
- Wer darf Bestellungen und Änderungen freigeben?
- Wer darf Störungsmeldungen im Namen des Kunden aufgeben?
- Welche Kommunikationswege sind gewünscht (Telefon, E-Mail, Ticketsystem)?

---

## Abschnitt 10: Bestandsdokumentation

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

## Abschluss

- [ ] Alle Abschnitte 1 bis 10 sind bearbeitet oder mit „unbekannt“ markiert.
- [ ] Keine Passwörter oder Geheimnisse in diesem Dokument eingetragen.
- [ ] Offene Punkte sind notiert und einem Verantwortlichen zugeordnet.
- [ ] Fragebogen mit dem Kunden besprochen.

**Offene Punkte:**

| Nr. | Thema | Verantwortlich | Fällig bis |
|---|---|---|---|
| | | | |
| | | | |

**Unterschrift Kunde:** ____________________________________ **Datum:** ____________

**Unterschrift Onboarding-Mitarbeiter:** ____________________________________ **Datum:** ____________
