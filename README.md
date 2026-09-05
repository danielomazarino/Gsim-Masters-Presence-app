# Träningsnärvaro Göteborg Sim

Enkel och snabb mobilapplikation för träningsnärvaro.

## Översikt
- **Source of Truth:** Google Sheet är master för all närvarodata och medlemshantering.
- **Miljö:** Körs som en standalone webbapplikation via Google Apps Script `HtmlService`.
- **Kommunikation:** Backendanrop sker uteslutande via `google.script.run`.
- **Aktuell status:** Steg 1, Steg 2, Steg 3 och Steg 4 implementerade.
  - Steg 1: Mobilanpassad snabb närvarovy (Idag / Igår / 2 dagar sedan, Askim / Valhalla, touchvänliga rader, diskret kontrollräkning).
  - Steg 2: Diskret kugghjul som navigerar till administration.
  - Steg 3: `saveAttendance(date, location, attendance)` för att spara närvaro och skapa saknade träningskolumner.
  - Steg 4: `deleteTrainingColumn(date, location)` i administrationen för att fysiskt radera felaktigt skapade träningskolumner med `deleteColumn()`, skyddat av `LockService` och efterföljande verifiering.
- **Filer:**
  - `Index.html`: Fristående Apps Script-fil med HTML, CSS och JavaScript i ett, redo att klistras in i Google Apps Script.
  - `Kod.gs`: Google Apps Script backendfunktioner inklusive `deleteTrainingColumn()`.
