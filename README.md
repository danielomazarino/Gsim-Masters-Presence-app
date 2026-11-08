# Träningsnärvaro (training presence app)

Enkel och snabb mobilapplikation för träningsnärvaro.

## Översikt
- **Source of Truth:** Google Sheet är master för all närvarodata och medlemshantering.
- **Miljö:** Körs som en standalone webbapplikation via Google Apps Script `HtmlService`.
- **Kommunikation:** Backendanrop sker uteslutande via `google.script.run`.
- **Aktuell status:** Steg 1, Steg 2, Steg 3 och Steg 4 implementerade.
  - Steg 1: Mobilanpassad snabb närvarovy (Idag / Igår / 2 dagar sedan, val av träningsgrupp, touchvänliga rader, diskret kontrollräkning).
  - Steg 2: Diskret kugghjul som navigerar till administration.
  - Steg 3: `saveAttendance(date, location, attendance)` för att spara närvaro och skapa saknade träningskolumner.
  - Steg 4: `deleteTrainingColumn(date, location)` i administrationen för att fysiskt radera felaktigt skapade träningskolumner med `deleteColumn()`, skyddat av `LockService` och efterföljande verifiering.
- **Filer:**
  - `Index.html`: Fristående Apps Script-fil med HTML, CSS och JavaScript i ett, redo att klistras in i Google Apps Script.
  - `Kod.gs`: Google Apps Script backendfunktioner inklusive `deleteTrainingColumn()`.
- **PWA:** `pwa/` innehåller en statisk wrapper (GitHub Pages) för hemskärmsikon på iPhone/Android. Ikonfiler i `icons/`.

## Start och utloggning

`pwa/index.html` är startpunkten för hemskärmsikonen. Wrappern visar ett startkort
med **Logga in** tills användaren har loggat in en gång, och går därefter direkt
till appen vid varje start.

```mermaid
sequenceDiagram
    participant Icon as Hemskärmsikon
    participant Wrapper as GitHub Pages (pwa/)
    participant App as Apps Script (exec)
    participant Google as Google-konton
    Icon->>Wrapper: start
    alt pwaSeen satt (inloggad tidigare)
        Wrapper->>App: location.replace
        App-->>Icon: appen direkt, ingen inloggning
    else inte inloggad än
        Wrapper-->>Icon: startkort med Logga in
        Icon->>Wrapper: Logga in (sätter pwaSeen)
        Wrapper->>App: APP_URL
        App->>Google: session saknas
        Google-->>App: inloggning
    end
```

Vid utloggning kan Google inte skicka tillbaka webbläsaren till GitHub-sidan:
`accounts.google.com/Logout` svarar **HTTP 400** för alla `continue`-URL:er som inte
ligger på en Google-domän. Därför avslutas sessionen via Workspace-utloggningen, som
i sin tur landar på appens egen inloggningssida.

```mermaid
sequenceDiagram
    participant App as Appen (kugghjul)
    participant Wrapper as GitHub Pages (pwa/)
    participant Google as Google-konton
    App->>Wrapper: ?logout=1 (rensar pwaSeen)
    Wrapper->>Google: www.google.com/a/goteborgsim.se/Logout?continue=<exec>
    Google->>Google: avslutar SSO- och Google-session
    Google->>App: redirect till exec-URL (utloggad)
    App->>Google: Workspace-inloggning
    Note over Google: nästa Logga in ger en helt ny session
```

Startflaggan `pwaSeen` sätts **endast** när användaren klickar på Logga in, aldrig
bara för att startkortet visas. iOS ger en hemskärmsapp en egen lagringsyta, så en
flagga satt i Safari skulle annars göra att första trycket på en ny ikon hamnade på
Googles inloggningssida.

