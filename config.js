// Konfiguration der Medikamenten-App.
// Nach der App-Registrierung in Microsoft Entra ID hier die beiden IDs eintragen.
window.APP_CONFIG = {
  // "Anwendungs-ID (Client)" aus der Entra-App-Registrierung
  clientId: "e99fbf79-07b1-425e-a12d-c1d65c6aaad5",

  // "Verzeichnis-ID (Mandant)" aus der Entra-App-Registrierung
  tenantId: "e4e57d17-93c0-4e0e-a004-ec2c9ea7a8e0",

  // Pfad der Excel-Datei in deinem OneDrive (relativ zu "Dokumente")
  filePath: "1 - Dirk/09 Gesundheit/Doku Herzinfarkt.xlsx",

  // Name des Tabellenblatts
  sheetName: "Inventur",

  // So viele Tage vor Ende des Vorrats liegt die Erinnerung
  vorlaufTage: 21,

  // Uhrzeit des Erinnerungstermins im Kalender
  erinnerungUhrzeit: "09:00",

  // Zeitzone für Kalendertermine (Windows-Bezeichnung)
  zeitzone: "W. Europe Standard Time"
};
