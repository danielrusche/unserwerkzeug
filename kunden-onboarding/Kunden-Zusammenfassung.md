# IT-Onboarding: Kunden-Zusammenfassung

**Stand:** 12.11.2026, 20:30 UTC
**Grundlage:** Kunden-Onboarding-Fragebogen-v2.md (Abschnitte 1 bis 11), Arbeitsmappe Kunden-Onboarding-Excel.xlsx
**Kunde:** offen
**Onboarding durchgeführt von:** offen
**Datum des Onboardings:** offen

## Stand der Erfassung

Der Fragebogen ist angelegt, aber noch nicht mit Kundenangaben befüllt. Es liegen keine Antworten vor. Die AD-Auslese (Nutzer, Computer, Standorte, Gruppen) wurde noch nicht ausgeführt; im Ordner gibt es deshalb keine CSV-Dateien. Die Tabs für Nutzer, Computer, Standorte und Gruppen aus der AD-Auslese sind daher noch nicht in der Mappe enthalten.

## Wichtigste Informationen

- Domäne (intern / öffentlich): offen
- Anzahl Standorte: offen
- Anzahl Nutzer insgesamt: offen
- Anzahl Clients und Notebooks: offen
- Anzahl Server und Hypervisor-Hosts: offen
- Virtualisierungsplattform: offen
- Firewall (Hersteller, Modell): offen
- Internetprovider und Anschlusstyp: offen
- Bestehende Namenskonvention: offen

## Ansprechpartner

| Rolle | Name | E-Mail | Telefon (Büro) | Telefon (mobil) | Verfügbarkeit |
|---|---|---|---|---|---|
| Hauptansprechpartner | offen | offen | offen | offen | offen |
| Stellvertretung | offen | offen | offen | offen | offen |
| Geschäftsführung / Entscheider | offen | offen | offen | offen | offen |
| Notfallkontakt (außerhalb der Geschäftszeiten) | offen | offen | offen | offen | offen |
| Externer Dienstleister (Telefonie, Internet, Drucker) | offen | offen | offen | offen | offen |

Hauptsitz-Adresse: offen

## Standorte

| Standort | Adresse | Ansprechpartner vor Ort | Nutzer | Internetanschluss |
|---|---|---|---|---|
| offen | offen | offen | offen | offen |

Die Standortliste steht im Tab „2 Standorte“. Serverraum oder Technikraum je Standort: offen.

## Nutzer und Systeme

- **Nutzer:** Die Erfassung ist im Tab „3 Nutzer“ vorbereitet. Anzahl gesamt: offen. Abteilungen und Leitungen: offen. Gäste oder externe Zugänge: offen.
- **Systeme:** Die Erfassung ist in den Tabs „5 Systeme“ und „8 Endgeräte“ vorbereitet. Server, Hypervisoren, Drucker, Clients: alle offen.
- **Administrative Zugänge:** Nur Benutzernamen und Ablageorte der Passwörter, keine Passwörter. Im Tab „4 Admin-Zugänge“ alle offen.

## Offene Punkte

Alle Abschnitte sind noch offen. Die wichtigsten Fragen für den Erstkontakt:

1. Name des Kunden, Hauptansprechpartner und Geschäftsführung mit Kontaktdaten: offen
2. Standorte mit Adressen, Nutzerzahlen und Internetanschlüssen: offen
3. Domäne, Domain Controller und Anzahl Nutzer und Computer: offen. Diese Angaben kommen aus der AD-Auslese, sobald sie ausgeführt ist.
4. Virtualisierung, Server und Management-Zugänge (iDRAC, iLO, IPMI): offen
5. Firewall, VPN und Internetanschlüsse mit Vertragslaufzeiten: offen
6. Vorhandene Dokumentation (Netzplan, Bestandsliste, Notfallhandbuch, Backup-Konzept) und ihr Ablageort: offen
7. Wer darf Bestellungen und Störungsmeldungen freigeben: offen

## Nächster Schritt

Die Kundenangaben in den gelben Feldern der Arbeitsmappe eintragen. Die AD-Erhebung (AD-Erhebung.ps1) auf einem Rechner mit Zugriff auf die Domäne ausführen; die CSV-Dateien kommen dann als zusätzliche Tabs in die Mappe. Passwörter gehören nicht in die Mappe, nur ihr Ablageort.
