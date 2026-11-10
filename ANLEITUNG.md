# Medikamente-App – Einrichtung

Web-App für das iPhone. Datenquelle bleibt die Excel-Datei
`1 - Dirk/09 Gesundheit/Doku Herzinfarkt.xlsx`, Blatt **Inventur**.

## Dateien

| Datei | Zweck |
|---|---|
| `index.html` | Oberfläche |
| `app.js` | Logik (Berechnung, Excel lesen/schreiben, Kalender) |
| `config.js` | Einstellungen (Client-ID, Tenant-ID, Dateipfad, Vorlauf 21 Tage) |
| `msal-browser.min.js` | Microsoft-Anmeldebibliothek (v3.28.1, lokal mitgeliefert) |
| `manifest.webmanifest`, `icon-180.png`, `icon-512.png` | App-Symbol für den Home-Bildschirm |

## 1. Webspace festlegen

Alle Dateien in **einen** Ordner auf deinem Webspace hochladen, z. B.
`https://www.maravia.de/medikamente/`. **HTTPS ist Pflicht.**

Die exakte Adresse merken – sie wird in Schritt 2 als Umleitungs-URI gebraucht.
Wichtig: Die App meldet sich mit der Adresse zurück, unter der sie aufgerufen wurde.
`.../medikamente/` und `.../medikamente/index.html` sind für Microsoft zwei verschiedene
Adressen. Am besten beide eintragen.

## 2. App-Registrierung in Microsoft Entra ID

1. https://entra.microsoft.com öffnen → **Identität → Anwendungen → App-Registrierungen → Neue Registrierung**
2. Name: `Medikamente-App`
3. Unterstützte Kontotypen: **Nur Konten in diesem Organisationsverzeichnis (nur maravia)**
4. Umleitungs-URI: Plattform **Single-Page-Anwendung (SPA)**, URI aus Schritt 1
   (z. B. `https://www.maravia.de/medikamente/`) → **Registrieren**
5. Unter **Authentifizierung** ggf. die zweite URI (`.../index.html`) ergänzen.
6. Unter **API-Berechtigungen → Berechtigung hinzufügen → Microsoft Graph → Delegierte Berechtigungen**:
   - `Files.ReadWrite`
   - `Calendars.ReadWrite`
   - (`User.Read` ist bereits vorhanden)
   
   Optional: **Administratorzustimmung für maravia erteilen**, dann entfällt die Rückfrage beim ersten Anmelden.
7. Auf der Seite **Übersicht** kopieren:
   - **Anwendungs-ID (Client)** → in `config.js` bei `clientId`
   - **Verzeichnis-ID (Mandant)** → in `config.js` bei `tenantId`

Danach `config.js` erneut hochladen.

## 3. Auf dem iPhone installieren

1. Adresse in **Safari** öffnen
2. **Teilen** → **Zum Home-Bildschirm**
3. App über das neue Symbol starten → **Mit Microsoft anmelden**

## 4. Kalender-Erinnerungen

Die App legt die Termine „Nachbestellen: …“ in deinem **Microsoft-365-Kalender** an
(21 Tage vor Ende des Vorrats, 09:00 Uhr, mit Erinnerung).
Damit sie auf dem iPhone erscheinen, muss der Microsoft-365-Kalender dort synchronisiert sein
(iOS-Einstellungen → Kalender → Accounts, oder die Outlook-App).

- Nach **Einlagern/Gezählt** wird die Erinnerung des betroffenen Medikaments automatisch neu gesetzt.
- Der Button **Kalender-Erinnerungen aktualisieren** setzt alle neu (beim ersten Mal einmal ausführen).
- Alte, von der App angelegte Termine werden dabei gelöscht und ersetzt; andere Termine bleiben unberührt.

## Bedienung

- **Übersicht:** Bestand heute, Reichweite in Tagen, „reicht bis“-Datum. Oben die Medikamente,
  die innerhalb von 21 Tagen ausgehen.
- **Antippen → Eingelagert:** Anzahl neu hinzugekommener Tabletten eingeben. Die App schreibt in Excel
  `Menge = heutiger Rest + Zugang` und `gezählt am = heute`. Deine Formeln (reicht bis, Anzahl noch übrig, Wert)
  rechnen unverändert weiter.
  Optional **Packungspreis** und **Stück pro Packung** eingeben: Die App berechnet den Preis pro Stück
  (Packungspreis / Stück) und schreibt ihn in Spalte I „Preis/Stck“. „Stück pro Packung“ wird pro Medikament
  auf dem Gerät gemerkt.
- **Antippen → Gezählt:** tatsächlich gezählten Bestand eintragen (Korrektur).
- **Ausblenden:** für Medikamente, die du nicht mehr nimmst (z. B. Prasugrel). Gilt nur für dieses Gerät;
  die Erinnerung wird dabei entfernt.
- Neue Medikamente oder geänderte Einnahme (morgens/abends) weiterhin direkt in Excel pflegen –
  die App übernimmt sie beim nächsten Aktualisieren.

## Demo

`https://<deine-adresse>/?demo` zeigt die App mit den Daten aus der Tabelle, ohne Anmeldung und ohne zu speichern.

## Einschränkungen

- Die Spaltenreihenfolge im Blatt Inventur (A Medikament … F gezählt am, I Preis/Stck, K Notiz) muss erhalten bleiben.
  Die App prüft die Kopfzeile und meldet Abweichungen.
- Microsoft-Anmeldungen im Browser bleiben in der Regel ca. 24 Stunden ohne Rückfrage gültig.
  Danach wird beim Öffnen ggf. kurz zur Microsoft-Anmeldung weitergeleitet.
