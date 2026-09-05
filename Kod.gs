/**
 * ==========================================================================
 * Google Apps Script - Träningsnärvaro Göteborg Sim
 * ==========================================================================
 * Denna fil innehåller backendfunktioner för Google Sheet och HtmlService.
 *
 * Existerande funktioner i Apps Script-backend (Source of Truth):
 * - doGet(e): Serverar Index.html
 * - getAttendanceData(): Läser medlemmar, kolumner och närvarostatus
 * - saveAttendance(date, location, attendance): Sparar närvaro och säkerställer
 *   träningskolumn vid behov
 *
 * STEP 4 FUNKTION:
 * - deleteTrainingColumn(date, location): Tar bort en specifik träningskolumn
 *   fysiskt med sheet.deleteColumn() skyddad av LockService, strikt dubblettskydd,
 *   A/B-systemkolumnskydd, regex-validering och efterföljande verifiering mot rad 1.
 * ==========================================================================
 */

/**
 * Normaliserar datum till formatet D/M (t.ex. "05/09" -> "5/9", "5/9" -> "5/9").
 * Hanterar både d/M och dd/MM.
 * @param {string} dateStr
 * @return {string}
 */
function normalizeTrainingDate(dateStr) {
  if (!dateStr || typeof dateStr !== 'string') return '';
  var trimmed = dateStr.trim();
  var parts = trimmed.split('/');
  if (parts.length === 2) {
    var day = parseInt(parts[0], 10);
    var month = parseInt(parts[1], 10);
    if (!isNaN(day) && !isNaN(month) && day >= 1 && day <= 31 && month >= 1 && month <= 12) {
      return day + '/' + month;
    }
  }
  return trimmed;
}

/**
 * Tar bort en felaktigt skapad träningskolumn fysiskt ur kalkylarket.
 *
 * Säkerhetskrav (Step 4 härdning):
 * 1. Tillåt endast 'Askim' eller 'Valhalla'.
 * 2. Normalisera datum d/M och dd/MM.
 * 3. Bygg exakt förväntad rubrik: `${location} ${normalizedDate}`.
 * 4. Läs rad 1 och samla ALLA kolumner vars rubrik exakt matchar den förväntade rubriken.
 * 5. Säkerhetsregel:
 *    - 0 matchningar -> returnera { success: false, deleted: false, ... }. Ingenting tas bort.
 *    - exakt 1 matchning -> fortsätt.
 *    - fler än 1 matchning -> kasta fel/avbryt med fel. Ingenting tas bort.
 * 6. Skydda systemkolumnerna A och B uttryckligen: om kolumn <= 2 avbryts operationen med fel.
 * 7. Kontrollera återigen att den hittade kolumnens faktiska rubrik exakt motsvarar förväntad rubrik.
 * 8. Kontrollera att rubriken är en giltig träningskolumn enligt /^(Askim|Valhalla)\s+\d{1,2}\/\d{1,2}$/.
 * 9. Använd LockService.getScriptLock() och releaseLock() i finally.
 * 10. Ta bort HELA kolumnen fysiskt med sheet.deleteColumn(foundColumn).
 * 11. SpreadsheetApp.flush() efter deletion.
 * 12. Verifiera efteråt att den exakta träningsrubriken inte längre finns i rad 1.
 *
 * @param {string} date - T.ex. "5/9" eller "05/09"
 * @param {string} location - "Askim" eller "Valhalla"
 * @return {Object} Resultatobjekt med status och detaljer
 */
function deleteTrainingColumn(date, location) {
  var lock = LockService.getScriptLock();
  try {
    // 9. Försök erhålla lås för att förhindra samtidiga strukturändringar i arket
    var hasLock = lock.tryLock(15000);
    if (!hasLock) {
      return {
        success: false,
        deleted: false,
        error: "Kunde inte erhålla scriptlås. En annan operation pågår. Försök igen om en stund."
      };
    }

    // 1. Validera location: Endast 'Askim' eller 'Valhalla'
    if (!location || (location !== 'Askim' && location !== 'Valhalla')) {
      return {
        success: false,
        deleted: false,
        error: "Ogiltig plats: '" + location + "'. Endast 'Askim' eller 'Valhalla' är tillåtna."
      };
    }

    // 2. Validera och normalisera datum
    if (!date || typeof date !== 'string') {
      return {
        success: false,
        deleted: false,
        error: "Ogiltigt eller saknat datum."
      };
    }

    var normalizedDate = normalizeTrainingDate(date);
    if (!normalizedDate) {
      return {
        success: false,
        deleted: false,
        error: "Ogiltigt datumformat: '" + date + "'. Förväntat format är D/M (t.ex. 5/9)."
      };
    }

    // 3. Konstruera exakt förväntad rubrik
    var expectedHeader = location + ' ' + normalizedDate;

    // 8. Kontrollera att rubriken är en giltig träningskolumn enligt regex
    var trainingHeaderPattern = /^(Askim|Valhalla)\s+\d{1,2}\/\d{1,2}$/;
    if (!trainingHeaderPattern.test(expectedHeader)) {
      return {
        success: false,
        deleted: false,
        date: normalizedDate,
        location: location,
        header: expectedHeader,
        error: "Rubriken '" + expectedHeader + "' matchar inte tillåtet format för träningskolumner."
      };
    }

    // Hämta aktivt kalkylblad
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      return {
        success: false,
        deleted: false,
        error: "Kunde inte öppna det aktiva kalkylarket."
      };
    }

    var sheet = ss.getSheetByName("HT -26");
    if (!sheet) {
      return {
        success: false,
        deleted: false,
        error: "Kunde inte hitta bladet 'HT -26' i kalkylarket."
      };
    }

    var lastColumn = sheet.getLastColumn();
    if (lastColumn < 1) {
      return {
        success: false,
        deleted: false,
        error: "Kalkylarket är tomt."
      };
    }

    // 4. Läs rad 1 och samla ALLA kolumner vars rubrik exakt matchar den förväntade rubriken
    var headerRowValues = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
    var matchedColumnIndices = [];

    for (var colIdx = 0; colIdx < headerRowValues.length; colIdx++) {
      var headerVal = String(headerRowValues[colIdx]).trim();
      if (headerVal === expectedHeader) {
        matchedColumnIndices.push(colIdx + 1); // 1-baserat kolumnindex i Google Sheets
      }
    }

    // 5. Säkerhetsregel för matchningar:
    // 0 matchningar -> returnera success: false, deleted: false. Ingenting får tas bort.
    if (matchedColumnIndices.length === 0) {
      return {
        success: false,
        deleted: false,
        date: normalizedDate,
        location: location,
        header: expectedHeader,
        error: "Hittade ingen kolumn med rubriken '" + expectedHeader + "'. Ingen ändring gjordes."
      };
    }

    // Fler än 1 matchning -> kasta fel och avbryt. Ingenting får tas bort.
    if (matchedColumnIndices.length > 1) {
      return {
        success: false,
        deleted: false,
        date: normalizedDate,
        location: location,
        header: expectedHeader,
        error: "Hittade " + matchedColumnIndices.length + " identiska kolumner med rubriken '" + expectedHeader + "'. Radering avbröts för att skydda dataintegriteten."
      };
    }

    // Exakt 1 matchning -> fortsätt
    var foundColumn = matchedColumnIndices[0];

    // 6. Skydda systemkolumnerna A och B uttryckligen: om kolumn <= 2 ska funktionen avbryta med fel
    if (foundColumn <= 2) {
      return {
        success: false,
        deleted: false,
        date: normalizedDate,
        location: location,
        header: expectedHeader,
        error: "Säkerhetsspärr: Kolumn " + foundColumn + " är en systemkolumn (A eller B) och får inte raderas."
      };
    }

    // 7. Kontrollera återigen att den hittade kolumnens faktiska rubrik exakt motsvarar den förväntade rubriken innan deletion
    var actualHeaderBeforeDelete = String(sheet.getRange(1, foundColumn).getValue()).trim();
    if (actualHeaderBeforeDelete !== expectedHeader) {
      return {
        success: false,
        deleted: false,
        date: normalizedDate,
        location: location,
        header: expectedHeader,
        error: "Avbröt: Kolumn " + foundColumn + " har rubriken '" + actualHeaderBeforeDelete + "', vilket inte exakt matchar förväntad rubrik '" + expectedHeader + "'."
      };
    }

    // 8. Kontrollera även den faktiska cellrubriken mot träningskolumn-regex innan radering
    if (!trainingHeaderPattern.test(actualHeaderBeforeDelete)) {
      return {
        success: false,
        deleted: false,
        date: normalizedDate,
        location: location,
        header: actualHeaderBeforeDelete,
        error: "Avbröt: Rubriken '" + actualHeaderBeforeDelete + "' är inte en giltig träningskolumn."
      };
    }

    // 10. Ta bort HELA kolumnen fysiskt med sheet.deleteColumn(foundColumn) (inte clearContent)
    sheet.deleteColumn(foundColumn);

    // 11. SpreadsheetApp.flush() efter deletion
    SpreadsheetApp.flush();

    // 12. Verifiera efteråt att den exakta träningsrubriken inte längre finns i rad 1
    var updatedLastCol = sheet.getLastColumn();
    var isStillPresent = false;
    if (updatedLastCol >= 1) {
      var updatedHeaders = sheet.getRange(1, 1, 1, updatedLastCol).getValues()[0];
      for (var u = 0; u < updatedHeaders.length; u++) {
        if (String(updatedHeaders[u]).trim() === expectedHeader) {
          isStillPresent = true;
          break;
        }
      }
    }

    if (isStillPresent) {
      return {
        success: false,
        deleted: false,
        date: normalizedDate,
        location: location,
        header: expectedHeader,
        deletedColumn: foundColumn,
        error: "Kunde inte verifiera radering av '" + expectedHeader + "'. Rubriken finns fortfarande kvar i rad 1."
      };
    }

    // 13. Returstruktur vid lyckad radering
    return {
      success: true,
      deleted: true,
      date: normalizedDate,
      location: location,
      header: expectedHeader,
      deletedColumn: foundColumn,
      message: "'" + expectedHeader + "' har raderats permanent från kalkylarket."
    };

  } catch (err) {
    return {
      success: false,
      deleted: false,
      error: "Fel vid radering av träningskolumn: " + (err.message || String(err))
    };
  } finally {
    // 9. Frigör alltid låset i finally
    try {
      lock.releaseLock();
    } catch (e) {
      // Tyst ignorering om låset redan släppts
    }
  }
}
