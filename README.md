# Träningsnärvaro Göteborg Sim

Enkel och snabb mobilapplikation för träningsnärvaro.

## Översikt
- **Source of Truth:** Google Sheet är master för all närvarodata och medlemshantering.
- **Miljö:** Körs som en standalone webbapplikation via Google Apps Script `HtmlService`.
- **Kommunikation:** Backendanrop sker uteslutande via `google.script.run`.
- **Aktuell status:** Versionen är helt read-only (Steg 1 och Steg 2 klara).
  - Steg 1: Mobilanpassad snabb närvarovy (Idag / Igår / 2 dagar sedan, Askim / Valhalla, touchvänliga rader, diskret kontrollräkning).
  - Steg 2: Diskret kugghjul som navigerar till administration (placeholder, inga skrivfunktioner).
  - Steg 3: `saveAttendance()` är ännu **inte** implementerat i frontend.
- **Filer:**
  - `Index.html`: Fristående Apps Script-fil med HTML, CSS och JavaScript i ett, redo att klistras in i Google Apps Script.
