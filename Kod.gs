// ============================================================
// TRÄNINGSNÄRVARO — GOOGLE APPS SCRIPT BACKEND
// ============================================================
//
// Spreadsheet: Träningsnärvaro HT -26
// Sheet:       HT -26
//
// Architecture:
//   Option B — Apps Script HtmlService + google.script.run
//
// Main functions:
//   A  getAttendanceData()
//   B  ensureTrainingColumn()
//   C  saveAttendance()
//   F0 deleteTrainingColumn()
//   P  resolveAttendanceSheet() / getPaymentStatusMap()
//   Q  setPaymentStatuses()
//
// Payment status lives in column A background colors:
//   #6aa84f (green) -> "paid",  #990000 (dark red) -> "owes",
//   anything else   -> "none"
// The colors can be administered BOTH from the web app and manually
// in Google Sheets — the sheet remains the single source of truth.
// The sheet itself is resolved robustly (see resolveAttendanceSheet),
// so a renamed sheet (e.g. "VT -27") or a new spreadsheet works
// without code changes.
// ============================================================

// ------------------------------------------------------------
// P — SHEET RESOLUTION & PAYMENT STATUS
// ------------------------------------------------------------

var PAYMENT_COLOR_PAID = "#6aa84f";
var PAYMENT_COLOR_OWES = "#990000";
var PAYMENT_COLOR_NONE = "#ffffff";

// Fallback sheet names tried in order before content-based detection.
// Extend this list for new seasons (e.g. add "VT -27", "HT -27").
var KNOWN_SHEET_NAMES = ["HT -26", "VT -27"];

/**
 * Resolves the attendance sheet robustly.
 *
 * Order:
 *   1. Exact-name match against KNOWN_SHEET_NAMES (fast path).
 *   2. Content-based detection: first sheet whose row 1 contains a
 *      header that looks like a training column
 *      (location + d/M pattern, e.g. "Askim 5/9").
 *   3. Fallback: the first sheet in the workbook.
 *
 * Survives sheet renames ("HT -26" -> "VT -27"), new season sheets,
 * and brand-new spreadsheet copies.
 *
 * @param {Spreadsheet} ss - The active spreadsheet.
 * @return {Sheet} The resolved attendance sheet.
 */
function resolveAttendanceSheet(ss) {
  // 1. Known names, fast path
  for (var i = 0; i < KNOWN_SHEET_NAMES.length; i++) {
    var byName = ss.getSheetByName(KNOWN_SHEET_NAMES[i]);
    if (byName) {
      return byName;
    }
  }

  // 2. Content-based detection
  var trainingHeaderPattern = /\s+\d{1,2}\/\d{1,2}$/;
  var sheets = ss.getSheets();
  for (var j = 0; j < sheets.length; j++) {
    var candidate = sheets[j];
    var cols = candidate.getLastColumn();
    if (cols < 1) {
      continue;
    }
    var headerRow = candidate
      .getRange(1, 1, 1, cols)
      .getValues()[0];
    for (var k = 0; k < headerRow.length; k++) {
      var header = String(headerRow[k]).trim();
      if (header && trainingHeaderPattern.test(header)) {
        return candidate;
      }
    }
  }

  // 3. Fallback: first sheet
  if (sheets.length > 0) {
    return sheets[0];
  }

  throw new Error("Kalkylarket innehåller inga flikar.");
}

/**
 * Reads member payment status from column A background colors.
 *
 * @param {Sheet} sheet - The resolved attendance sheet.
 * @return {Object} Map: { "Member Name": "paid" | "owes" | "none" }.
 */
function getPaymentStatusMap(sheet) {
  var lastRow = sheet.getLastRow();
  var statusMap = {};

  if (lastRow < 2) {
    return statusMap;
  }

  var names = sheet
    .getRange(2, 1, lastRow - 1, 1)
    .getValues();
  var backgrounds = sheet
    .getRange(2, 1, lastRow - 1, 1)
    .getBackgrounds();

  for (var i = 0; i < names.length; i++) {
    var name = String(names[i][0]).trim();
    if (!name) {
      continue;
    }

    var bg = String(backgrounds[i][0]).trim().toLowerCase();

    if (bg === PAYMENT_COLOR_PAID) {
      statusMap[name] = "paid";
    } else if (bg === PAYMENT_COLOR_OWES) {
      statusMap[name] = "owes";
    } else {
      statusMap[name] = "none";
    }
  }

  return statusMap;
}

// ------------------------------------------------------------
// Q — WRITE PAYMENT STATUSES
// ------------------------------------------------------------

/**
 * Writes payment status changes to column A background colors.
 *
 * Coexists with manual editing in Google Sheets: the lock only
 * serializes app-driven writes, and every change is re-read and
 * verified after writing.
 *
 * @param {Array<Object>} changes - [{ name: string, status: "paid" | "owes" | "none" }].
 * @return {Object} { success: boolean, updated: Array, errors: Array }.
 */
function setPaymentStatuses(changes) {

  if (!Array.isArray(changes) || changes.length === 0) {
    return {
      success: false,
      updated: [],
      errors: ["Inga ändringar angivna."]
    };
  }

  var validStatuses = ["paid", "owes", "none"];
  for (var v = 0; v < changes.length; v++) {
    if (!changes[v] || typeof changes[v].name !== "string" || validStatuses.indexOf(changes[v].status) === -1) {
      return {
        success: false,
        updated: [],
        errors: ["Ogiltig ändring vid index " + v + "."]
      };
    }
  }

  var lock = LockService.getScriptLock();

  try {
    var hasLock = lock.tryLock(15000);

    if (!hasLock) {
      return {
        success: false,
        updated: [],
        errors: ["Kunde inte erhålla scriptlås. En annan operation pågår. Försök igen om en stund."]
      };
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = resolveAttendanceSheet(ss);

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return {
        success: false,
        updated: [],
        errors: ["Kalkylarket innehåller inga medlemmar."]
      };
    }

    // Fresh read right before writing: names + current backgrounds.
    // This minimizes the window for races against manual edits.
    var namesRange = sheet.getRange(2, 1, lastRow - 1, 1);
    var names = namesRange.getValues();
    var currentBackgrounds = namesRange.getBackgrounds();

    // Build row index: name -> array of row offsets (duplicates rejected)
    var nameToOffsets = {};
    for (var i = 0; i < names.length; i++) {
      var name = String(names[i][0]).trim();
      if (!name) {
        continue;
      }
      if (!nameToOffsets[name]) {
        nameToOffsets[name] = [];
      }
      nameToOffsets[name].push(i);
    }

    var colorFor = function(status) {
      if (status === "paid") return PAYMENT_COLOR_PAID;
      if (status === "owes") return PAYMENT_COLOR_OWES;
      return PAYMENT_COLOR_NONE;
    };

    // Stage writes: batch all changes into one setBackgrounds call.
    var staged = names.map(function() { return null; });
    var updated = [];
    var errors = [];

    changes.forEach(function(change) {
      var offsets = nameToOffsets[change.name];

      if (!offsets || offsets.length === 0) {
        errors.push("Medlemmen \"" + change.name + "\" hittades inte.");
        return;
      }
      if (offsets.length > 1) {
        errors.push("Medlemmen \"" + change.name + "\" finns på " + offsets.length + " rader — ändringen avbröts.");
        return;
      }

      var offset = offsets[0];
      var desiredColor = colorFor(change.status);
      var currentColor = String(currentBackgrounds[offset][0]).trim().toLowerCase();

      if (currentColor === desiredColor) {
        // Already correct (possibly edited manually in the meantime) — no write needed
        updated.push({ name: change.name, status: change.status, row: offset + 2, alreadySet: true });
        return;
      }

      staged[offset] = desiredColor;
      updated.push({ name: change.name, status: change.status, row: offset + 2 });
    });

    // Write all staged colors in one batch
    var anyWrites = staged.some(function(c) { return c !== null; });
    if (anyWrites) {
      var finalBackgrounds = currentBackgrounds.map(function(rowColors, idx) {
        return [staged[idx] !== null ? staged[idx] : rowColors[0]];
      });
      namesRange.setBackgrounds(finalBackgrounds);
      SpreadsheetApp.flush();

      // Re-read and verify every intended change
      var verifyBackgrounds = namesRange.getBackgrounds();
      updated.forEach(function(entry) {
        if (entry.alreadySet) {
          return;
        }
        var offset = entry.row - 2;
        var actual = String(verifyBackgrounds[offset][0]).trim().toLowerCase();
        var expected = colorFor(entry.status);
        if (actual !== expected) {
          errors.push("Kunde inte verifiera ändringen för \"" + entry.name + "\".");
        }
      });
    }

    return {
      success: errors.length === 0,
      updated: updated,
      errors: errors
    };

  } catch (err) {
    return {
      success: false,
      updated: [],
      errors: ["Fel vid sparande av betalstatus: " + (err.message || String(err))]
    };
  } finally {
    try {
      lock.releaseLock();
    } catch (e) {
      // Låset kan redan ha släppts
    }
  }
}


// ============================================================
// A — GET ATTENDANCE DATA
// ============================================================

function getAttendanceData() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = resolveAttendanceSheet(ss);
  const sheetName = sheet.getName();

  const spreadsheetName = ss.getName();
  const spreadsheetTimeZone = ss.getSpreadsheetTimeZone();

  const today = new Date();

  const todayDate =
    Utilities.formatDate(
      today,
      spreadsheetTimeZone,
      "d/M"
    );

  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  if (lastColumn < 1) {
    throw new Error(
      "Kunde inte läsa kalkylbladets kolumner."
    );
  }

  const values =
    sheet
      .getRange(
        1,
        1,
        Math.max(lastRow, 1),
        lastColumn
      )
      .getValues();

  const headers = values[0];

  // ----------------------------------------------------------
  // Hitta alla träningskolumner
  // ----------------------------------------------------------

  const trainingColumns = [];

  for (let i = 0; i < headers.length; i++) {

    const header =
      String(headers[i]).trim();

    const match =
      header.match(
        /^(.*?)\s+(\d{1,2}\/\d{1,2})$/
      );

    if (!match) {
      continue;
    }

    const location =
      match[1].trim();

    const date =
      match[2];

    trainingColumns.push({
      column: i + 1,
      columnLetter:
        columnNumberToLetter(i + 1),
      header: header,
      location: location,
      date: date
    });
  }

  // ----------------------------------------------------------
  // Träningar för dagens datum
  // ----------------------------------------------------------

  const todayTrainings =
    trainingColumns.filter(
      function(training) {
        return training.date === todayDate;
      }
    );

  let todayStatus = "NONE";

  if (todayTrainings.length === 1) {
    todayStatus = "ONE";
  } else if (todayTrainings.length > 1) {
    todayStatus = "MULTIPLE";
  }

  // ----------------------------------------------------------
  // Kontrollera dubletter
  // ----------------------------------------------------------

  const duplicateTrainings = [];
  const trainingKeyCounts = {};

  trainingColumns.forEach(
    function(training) {

      const key =
        training.location +
        "|" +
        training.date;

      if (!trainingKeyCounts[key]) {
        trainingKeyCounts[key] = [];
      }

      trainingKeyCounts[key].push(training);
    }
  );

  Object.keys(trainingKeyCounts).forEach(
    function(key) {

      const entries =
        trainingKeyCounts[key];

      if (entries.length > 1) {

        duplicateTrainings.push({
          location:
            entries[0].location,

          date:
            entries[0].date,

          columns:
            entries.map(
              function(entry) {
                return entry.column;
              }
            ),

          columnLetters:
            entries.map(
              function(entry) {
                return entry.columnLetter;
              }
            ),

          headers:
            entries.map(
              function(entry) {
                return entry.header;
              }
            )
        });
      }
    }
  );

  // ----------------------------------------------------------
  // Medlemslista
  // ----------------------------------------------------------

  const members = [];

  for (
    let row = 1;
    row < values.length;
    row++
  ) {

    const name =
      String(values[row][0]).trim();

    if (!name) {
      continue;
    }

    members.push({
      row: row + 1,
      name: name
    });
  }

  // ----------------------------------------------------------
  // Närvarodata
  // ----------------------------------------------------------

  const attendance = {};

  members.forEach(
    function(member) {

      attendance[member.name] = {};

      trainingColumns.forEach(
        function(training) {

          const cellValue =
            values[
              member.row - 1
            ][
              training.column - 1
            ];

          attendance[
            member.name
          ][
            training.header
          ] =
            String(cellValue)
              .trim()
              .toLowerCase() === "x";
        }
      );
    }
  );

  return {

    spreadsheetName:
      spreadsheetName,

    sheetName:
      sheetName,

    sheetUrl:
      ss.getUrl(),

    spreadsheetTimeZone:
      spreadsheetTimeZone,

    today:
      todayDate,

    todayStatus:
      todayStatus,

    trainingColumns:
      trainingColumns,

    todayTrainings:
      todayTrainings,

    duplicateTrainings:
      duplicateTrainings,

    members:
      members,

    attendance:
      attendance,

    paymentStatus:
      getPaymentStatusMap(sheet)
  };
}


// ------------------------------------------------------------
// TEST A — READ ONLY
// ------------------------------------------------------------

function testGetAttendanceData() {

  Logger.log(
    "=== GET ATTENDANCE DATA TEST ==="
  );

  const data =
    getAttendanceData();

  Logger.log(
    "Spreadsheet: " +
    data.spreadsheetName
  );

  Logger.log(
    "Sheet: " +
    data.sheetName
  );

  Logger.log(
    "Timezone: " +
    data.spreadsheetTimeZone
  );

  Logger.log(
    "Today: " +
    data.today
  );

  Logger.log(
    "Today's training status: " +
    data.todayStatus
  );

  Logger.log("");
  Logger.log(
    "--- TRAINING COLUMNS ---"
  );

  data.trainingColumns.forEach(
    function(training) {

      Logger.log(
        training.columnLetter +
        " | " +
        training.header +
        " | location=" +
        training.location +
        " | date=" +
        training.date
      );
    }
  );

  Logger.log("");
  Logger.log(
    "--- TODAY'S TRAININGS ---"
  );

  if (
    data.todayTrainings.length === 0
  ) {

    Logger.log("None");

  } else {

    data.todayTrainings.forEach(
      function(training) {

        Logger.log(
          training.columnLetter +
          " | " +
          training.header
        );
      }
    );
  }

  Logger.log("");
  Logger.log(
    "--- DUPLICATES ---"
  );

  if (
    data.duplicateTrainings.length === 0
  ) {

    Logger.log("None");

  } else {

    data.duplicateTrainings.forEach(
      function(duplicate) {

        Logger.log(
          duplicate.location +
          " " +
          duplicate.date +
          " | columns=" +
          duplicate.columnLetters.join(", ")
        );
      }
    );
  }

  Logger.log("");
  Logger.log(
    "--- MEMBERS ---"
  );

  Logger.log(
    "Member count: " +
    data.members.length
  );

  data.members
    .slice(0, 10)
    .forEach(
      function(member) {

        Logger.log(
          "Row " +
          member.row +
          ": " +
          member.name
        );
      }
    );

  Logger.log("");
  Logger.log(
    "=== TEST COMPLETE - NO DATA WAS CHANGED ==="
  );
}


// ============================================================
// B — ENSURE TRAINING COLUMN
// ============================================================

function ensureTrainingColumn(
  date,
  location
) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = resolveAttendanceSheet(ss);

  const allowedLocations = [
    "Askim",
    "Valhalla"
  ];

  if (!date || !location) {

    throw new Error(
      "Datum och lokal måste anges."
    );
  }

  if (
    !allowedLocations.includes(location)
  ) {

    throw new Error(
      'Ogiltig lokal: "' +
      location +
      '". Tillåtna lokaler är Askim eller Valhalla.'
    );
  }

  const dateMatch =
    String(date).match(
      /^(\d{1,2})\/(\d{1,2})$/
    );

  if (!dateMatch) {

    throw new Error(
      'Ogiltigt datumformat: "' +
      date +
      '". Använd formatet d/M, exempelvis 4/9.'
    );
  }

  const day =
    Number(dateMatch[1]);

  const month =
    Number(dateMatch[2]);

  if (
    day < 1 ||
    day > 31 ||
    month < 1 ||
    month > 12
  ) {

    throw new Error(
      "Ogiltigt datum."
    );
  }

  const normalizedDate =
    day +
    "/" +
    month;

  const requestedHeader =
    location +
    " " +
    normalizedDate;

  const lastColumn =
    sheet.getLastColumn();

  if (lastColumn < 1) {

    throw new Error(
      "Kunde inte läsa kalkylbladets kolumner."
    );
  }

  const headers =
    sheet
      .getRange(
        1,
        1,
        1,
        lastColumn
      )
      .getValues()[0];

  // ----------------------------------------------------------
  // Leta efter exakt befintlig kombination
  // ----------------------------------------------------------

  let foundColumn = null;

  for (
    let i = 0;
    i < headers.length;
    i++
  ) {

    const header =
      String(headers[i]).trim();

    if (
      header === requestedHeader
    ) {

      foundColumn =
        i + 1;

      break;
    }
  }

  // ----------------------------------------------------------
  // Befintlig kolumn
  // ----------------------------------------------------------

  if (foundColumn !== null) {

    return {

      success: true,

      created: false,

      date:
        normalizedDate,

      location:
        location,

      header:
        requestedHeader,

      column:
        foundColumn,

      columnLetter:
        columnNumberToLetter(
          foundColumn
        ),

      message:
        "Befintlig träningskolumn används."
    };
  }

  // ----------------------------------------------------------
  // Skapa ny fysisk kolumn
  // ----------------------------------------------------------

  const newColumn =
    sheet.getLastColumn() + 1;

  sheet.insertColumnAfter(
    sheet.getLastColumn()
  );

  sheet
    .getRange(
      1,
      newColumn
    )
    .setValue(
      requestedHeader
    );

  // ----------------------------------------------------------
  // Verifiera rubriken
  // ----------------------------------------------------------

  const verifiedHeader =
    String(
      sheet
        .getRange(
          1,
          newColumn
        )
        .getValue()
    ).trim();

  if (
    verifiedHeader !==
    requestedHeader
  ) {

    throw new Error(
      'Kolumnen skapades men rubriken kunde inte verifieras. ' +
      'Förväntade "' +
      requestedHeader +
      '", fick "' +
      verifiedHeader +
      '".'
    );
  }

  return {

    success: true,

    created: true,

    date:
      normalizedDate,

    location:
      location,

    header:
      requestedHeader,

    column:
      newColumn,

    columnLetter:
      columnNumberToLetter(
        newColumn
      ),

    message:
      "Ny träningskolumn skapad."
  };
}


// ------------------------------------------------------------
// TEST B — READ/IDEMPOTENCY TEST
// ------------------------------------------------------------
//
// This test uses the existing Askim 4/9 column.
// It does NOT create or delete anything.
// ------------------------------------------------------------

function testEnsureTrainingColumn() {

  Logger.log(
    "=== B — ENSURE TRAINING COLUMN TEST ==="
  );

  const result =
    ensureTrainingColumn(
      "4/9",
      "Askim"
    );

  Logger.log(
    "Success: " +
    result.success
  );

  Logger.log(
    "Created: " +
    result.created
  );

  Logger.log(
    "Header: " +
    result.header
  );

  Logger.log(
    "Column: " +
    result.columnLetter
  );

  if (
    result.success !== true
  ) {

    throw new Error(
      "B TEST MISSLYCKADES: success var inte true."
    );
  }

  if (
    result.header !==
    "Askim 4/9"
  ) {

    throw new Error(
      'B TEST MISSLYCKADES: förväntade "Askim 4/9".'
    );
  }

  if (
    result.created !== false
  ) {

    throw new Error(
      "B TEST MISSLYCKADES: Askim 4/9 skulle redan finnas."
    );
  }

  Logger.log("");
  Logger.log(
    "=== B TEST PASSED ==="
  );

  Logger.log(
    "Ingen data ändrades."
  );
}


// ============================================================
// C — SAVE ATTENDANCE
// ============================================================

function saveAttendance(
  date,
  location,
  attendance
) {

  if (
    !Array.isArray(attendance)
  ) {

    throw new Error(
      "attendance måste vara en array."
    );
  }

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    resolveAttendanceSheet(ss);

  // ----------------------------------------------------------
  // Säkerställ rätt träningskolumn
  // ----------------------------------------------------------

  const training =
    ensureTrainingColumn(
      date,
      location
    );

  const targetColumn =
    training.column;

  const targetColumnLetter =
    training.columnLetter;

  // ----------------------------------------------------------
  // Läs medlemslistan
  // ----------------------------------------------------------

  const lastRow =
    sheet.getLastRow();

  if (lastRow < 2) {

    throw new Error(
      "Det finns inga medlemsrader i kalkylbladet."
    );
  }

  const memberNames =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        1
      )
      .getValues();

  const validatedAttendance = [];
  const seenRows = {};

  // ----------------------------------------------------------
  // Validera inkommande närvaro
  // ----------------------------------------------------------

  for (
    let i = 0;
    i < attendance.length;
    i++
  ) {

    const item =
      attendance[i];

    if (
      !item ||
      typeof item.row !== "number"
    ) {

      throw new Error(
        "Ogiltig medlemsrad i attendance på position " +
        i +
        "."
      );
    }

    const row =
      item.row;

    if (
      row < 2 ||
      row > lastRow
    ) {

      throw new Error(
        "Medlemsrad " +
        row +
        " ligger utanför medlemslistan."
      );
    }

    if (
      seenRows[row]
    ) {

      throw new Error(
        "Medlemsrad " +
        row +
        " finns flera gånger i samma anrop."
      );
    }

    seenRows[row] = true;

    const name =
      String(
        memberNames[row - 2][0]
      ).trim();

    if (!name) {

      throw new Error(
        "Rad " +
        row +
        " innehåller inget medlemsnamn i kolumn A."
      );
    }

    if (
      typeof item.present !==
      "boolean"
    ) {

      throw new Error(
        'Närvarostatus för "' +
        name +
        '" måste vara true eller false.'
      );
    }

    validatedAttendance.push({

      row:
        row,

      name:
        name,

      present:
        item.present
    });
  }

  // ----------------------------------------------------------
  // Skriv närvaro
  // ----------------------------------------------------------

  const cellsToWrite = [];

  validatedAttendance.forEach(
    function(item) {

      const value =
        item.present
          ? "x"
          : "";

      sheet
        .getRange(
          item.row,
          targetColumn
        )
        .setValue(
          value
        );

      cellsToWrite.push({

        row:
          item.row,

        name:
          item.name,

        value:
          value
      });
    }
  );

  SpreadsheetApp.flush();

  // ----------------------------------------------------------
  // Verifiera skrivningen
  // ----------------------------------------------------------

  const verificationErrors = [];

  cellsToWrite.forEach(
    function(item) {

      const actual =
        String(
          sheet
            .getRange(
              item.row,
              targetColumn
            )
            .getValue()
        ).trim();

      const expected =
        String(
          item.value
        ).trim();

      if (
        actual !== expected
      ) {

        verificationErrors.push(
          item.name +
          " (rad " +
          item.row +
          "): förväntade '" +
          expected +
          "', fick '" +
          actual +
          "'"
        );
      }
    }
  );

  if (
    verificationErrors.length > 0
  ) {

    throw new Error(
      "Närvaroskrivningen kunde inte verifieras:\n" +
      verificationErrors.join("\n")
    );
  }

  // ----------------------------------------------------------
  // Resultat
  // ----------------------------------------------------------

  let presentCount = 0;
  let absentCount = 0;

  validatedAttendance.forEach(
    function(item) {

      if (item.present) {
        presentCount++;
      } else {
        absentCount++;
      }
    }
  );

  return {

    success:
      true,

    date:
      training.date,

    location:
      training.location,

    header:
      training.header,

    column:
      targetColumn,

    columnLetter:
      targetColumnLetter,

    createdColumn:
      training.created,

    updatedMembers:
      validatedAttendance.length,

    presentCount:
      presentCount,

    absentCount:
      absentCount,

    members:
      validatedAttendance.map(
        function(item) {

          return {

            row:
              item.row,

            name:
              item.name,

            present:
              item.present
          };
        }
      ),

    message:
      "Närvaro sparad och verifierad."
  };
}


// ------------------------------------------------------------
// TEST C — CONTROLLED WRITE + RESTORE
// ------------------------------------------------------------

function testSaveAttendance() {

  Logger.log(
    "=== C — SAVE ATTENDANCE TEST ==="
  );

  const date =
    "4/9";

  const location =
    "Askim";

  const christerRow =
    2;

  const amandaRow =
    3;

  // ----------------------------------------------------------
  // Säkerställ träningskolumn
  // ----------------------------------------------------------

  const training =
    ensureTrainingColumn(
      date,
      location
    );

  Logger.log("");
  Logger.log(
    "--- TRAINING ---"
  );

  Logger.log(
    "Header: " +
    training.header
  );

  Logger.log(
    "Column: " +
    training.columnLetter
  );

  Logger.log(
    "Created: " +
    training.created
  );

  // ----------------------------------------------------------
  // Hämta sheet
  // ----------------------------------------------------------

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(
        "HT -26"
      );

  // ----------------------------------------------------------
  // Läs ursprungliga värden
  // ----------------------------------------------------------

  const originalChrister =
    sheet
      .getRange(
        christerRow,
        training.column
      )
      .getValue();

  const originalAmanda =
    sheet
      .getRange(
        amandaRow,
        training.column
      )
      .getValue();

  Logger.log("");
  Logger.log(
    "--- ORIGINAL VALUES ---"
  );

  Logger.log(
    "Christer: '" +
    originalChrister +
    "'"
  );

  Logger.log(
    "Amanda: '" +
    originalAmanda +
    "'"
  );

  // ----------------------------------------------------------
  // Testa skrivning
  // ----------------------------------------------------------

  const result =
    saveAttendance(
      date,
      location,
      [
        {
          row:
            christerRow,

          present:
            true
        },

        {
          row:
            amandaRow,

          present:
            false
        }
      ]
    );

  Logger.log("");
  Logger.log(
    "--- WRITE RESULT ---"
  );

  Logger.log(
    "Success: " +
    result.success
  );

  Logger.log(
    "Header: " +
    result.header
  );

  Logger.log(
    "Column: " +
    result.columnLetter
  );

  Logger.log(
    "Updated members: " +
    result.updatedMembers
  );

  Logger.log(
    "Present: " +
    result.presentCount
  );

  Logger.log(
    "Absent: " +
    result.absentCount
  );

  // ----------------------------------------------------------
  // Kontrollera resultat
  // ----------------------------------------------------------

  const writtenChrister =
    sheet
      .getRange(
        christerRow,
        training.column
      )
      .getValue();

  const writtenAmanda =
    sheet
      .getRange(
        amandaRow,
        training.column
      )
      .getValue();

  if (
    String(writtenChrister).trim() !== "x"
  ) {

    throw new Error(
      "TEST MISSLYCKADES: Christer skulle ha x."
    );
  }

  if (
    String(writtenAmanda).trim() !== ""
  ) {

    throw new Error(
      "TEST MISSLYCKADES: Amanda skulle vara tom."
    );
  }

  // ----------------------------------------------------------
  // Återställ
  // ----------------------------------------------------------

  sheet
    .getRange(
      christerRow,
      training.column
    )
    .setValue(
      originalChrister
    );

  sheet
    .getRange(
      amandaRow,
      training.column
    )
    .setValue(
      originalAmanda
    );

  SpreadsheetApp.flush();

  // ----------------------------------------------------------
  // Verifiera återställning
  // ----------------------------------------------------------

  const restoredChrister =
    sheet
      .getRange(
        christerRow,
        training.column
      )
      .getValue();

  const restoredAmanda =
    sheet
      .getRange(
        amandaRow,
        training.column
      )
      .getValue();

  if (
    String(restoredChrister) !==
    String(originalChrister)
  ) {

    throw new Error(
      "Återställning av Christer misslyckades."
    );
  }

  if (
    String(restoredAmanda) !==
    String(originalAmanda)
  ) {

    throw new Error(
      "Återställning av Amanda misslyckades."
    );
  }

  Logger.log("");
  Logger.log(
    "=== C TEST PASSED — VALUES RESTORED ==="
  );
}


// ============================================================
// GEMENSAM HJÄLPFUNKTION
// ============================================================

function columnNumberToLetter(
  column
) {

  let letter = "";

  while (column > 0) {

    const remainder =
      (column - 1) % 26;

    letter =
      String.fromCharCode(
        65 + remainder
      ) +
      letter;

    column =
      Math.floor(
        (column - 1) / 26
      );
  }

  return letter;
}


// ============================================================
// D — OPTION B WEB APP
// ============================================================
//
// The web app now serves Index.html through HtmlService.
//
// The frontend communicates with the server using:
//   google.script.run
//
// This avoids the CORS problem encountered with the external
// AI Studio frontend calling ContentService directly.
//
// PWA icons (& manifest) are injected by doGet() as absolute
// links pointing at this same web-app URL (?pwa=<name>) — the
// only reliable way for iOS/Android to reach them, because the
// app itself runs in an iframe and cannot compute the exec URL.
// ============================================================

// ---- PWA ICONS (generated from /icons/ in the repo) --------
// Base64-encoded PNGs served as ?pwa=<name> from this web app,
// so the browser sees same-domain absolute icon URLs.
// Names must match what doGet() and the manifest reference.
var PWA_ICON_BASE64 = {
  "icon-192.png": "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",
  "icon-512.png": "iVBORw0KGgoAAAANSUhEUgAAAgAAAAIACAIAAAB7GkOtAAAuGklEQVR4nO3dd5wV1d3H8TN3Zm7Ze7fRFljKwsLSe1MQC4oilmiI0RiMNWpMfFJM7/EJ6U80GlNNTE+Mjz0WxIoNQUR6lb5LWWD73jbl+QMeKW7fuztn7vm8X758kbwS9uy9M7/vnDOnaMNvrBQAAPUEvG4AAMAbBAAAKIoAAABFEQAAoCgCAAAURQAAgKIIAABQFAEAAIoiAABAUQQAACiKAAAARREAAKAoAgAAFEUAAICiCAAAUBQBAACKIgAAQFEEAAAoigAAAEURAACgKAIAABRFAACAoggAAFAUAQAAiiIAAEBRBAAAKIoAAABFEQAAoCgCAAAURQAAgKIIAABQFAEAAIoiAABAUQQAACiKAAAARREAAKAoAgAAFEUAAICiCAAAUBQBAACKIgAAQFEEAAAoigAAAEURAACgKAIAABRFAACAoggAAFAUAQAAiiIAAEBRBAAAKIoAAABFEQAAoCgCAAAURQAAgKIIAABQFAEAAIoiAABAUQQAACiKAAAARREAAKAoAgAAFEUAAICiCAAAUBQBAACKIgAAQFEEAAAoigAAAEURAACgKAIAABRFAACAoggAAFAUAQAAiiIAAEBRBAAAKIoAAABFEQAAoCgCAAAURQAAgKIIAABQFAEAAIoiAABAUQQAACiKAAAARREAAKAoAgAAFEUAAICiCAAAUBQBAACKIgAAQFEEAAAoigAAAEURAACgKAIAABRFAACAoggAAFAUAQAAiiIAAEBRBAAAKIoAAABFEQAAoCgCAAAURQAAgKIMrxuA7qNpond+oG9hoFdeoGdeoGdeoFeeVhgLRMNaNKzFwtrRP5i6puvC0IUR0HRduK5IW27aFmnLTVvCst20LVKWSKXdmka3psGpbnj/325Ng1NV7x6osvdXObbj9S+M/2caorin3qcg0Cvv6D9az7xA7/xAYSwQMkXI1IKGCJpa8OgfDE0IYTuuZYu0LWzbtWxhOSKZcusTbn3crY+7DYmjf3Zq4+6hWudIrXuo1jlU6xypcyzb698WbUYAZK3e+YHhxXpZsVFSpA/oFRjYSy/uFTh6b7ePJvSgFj76pzazHXGw2qk4bJcfdioO2xVHnIrDznv77YrDtuu2uwlol3BQK+2nl/bTh/XTS/sZw/rrg/roejt7+3pACx4rD+343l1X1Da6lTVOxRG74rBTfvjov53yw3ZltePw1UtGG35jpddtQGYU99InDDEmDDXGDjbKio38aPtrfddrSLjbKuwt5dbWo/8utytr6ClkQCyiTR5mTiszp5eZ40oMQ/e6QR9g2WJPpb19v71jv73jgL19v719n11Vz7fvJQLAxzRNlBUbM0eZM0aaE4YaPXN9+Uanut5dvSO96j3rnW3p1TuseJKnxHaYMNS4YHLotFHmqIFGe5/xZVDT4L6339q0x960x9q0x9pcbnMBdCcCwH8KY4FzJgRnjzFPG2X6tOg3x3bE5r3W0TBYuTVdcYTHw6aNHWxcOC00f2qwuJd8j/qd4Lhi90F7015r4257/W5rzY50dT150IUIAN8Y0EufOyl43qTg5GGmH5/1OmD7fnvp2tSr69LLt6STaQqBKIhpV58d+fCs0KDeWVX3W7DroL1mh7V6u7V6R3rjbjtlcRlkEgEgu7wcbf600GUzQ5NLTa/b4plEyl2+Ob10XXrputTOAyrOMikp0q+fG7l8ZigclPHVTvdIW2LjHmvltvSKLem3t9I5yAACQFKaJmaOMq88KzxnQrAjU3ey19YK++nlyadWJBVJgsml5icvjMyZENS4Ck7gumJrhb1iS/roP0wl6BgCQDrRsHb5zNDCOZGhfVXp5nfM+l3WUyuSz6xIlh/Ozpu/f4/Al6+Izp8W8rohPvDi6tSt99Z63Qr/YR2ARHrnB268IHLlmeFomIe91o0ZbIwZbHxpQfTd7dbjbyYeX5ZsSGTJmEA4qH1yXuST8yIqD/i0S58CNV6LZRoBIIV+PQKfnJdzxexQyOSGbx9NE5NKjUmlsS9fEX3yreQ/X05s2G153ahOOWd88DsLY/17UNHQ5QgAj/XIDXzmkpwrzwybfBWdkxPSrjwzfOWZ4TU7rH++nHhqRTKR8lmHwDTElxZEr5sb8bohUAVVxzMhU7vuvPAt83NiEZ76M2n8EGP8kNjXrox+/Cc1m/f6pjcwqLd+9625YwdzS6L7cLV544Ipoa9fGe1HN7/L5OVoBVJuhtGk+dNCi66N8e4H3YwA6G79ewS+szB2zvig1w2BLG69KOcLl+d43QqoiADoPgFNXDs38rkP5URCPOhBCCE0TXz347GPnR32uiFQFAHQTfr3CPzkxtzpI9RdzYtTaJpYdG3sI2dQ/eEZAqA7XDIj9N2FsVxe9uIE/30N1R8eIwC6VsjUvrswumAW9zlO8tnLcj56JlcFPEYAdKHinoFf3pY3hol9ONmCWeFPX8xbX3iP2tRVZo0277o5ryDGsA9OMrbE+N41Ua9bAQhBAHSRK2aH77wmpsiu/Wi7vBztl7flsb0rJEEAZN5nL8uhg48mfY9NfiATAiCT9ID4wXW5l89k/1404cKpoYumc21AIgRAxugBcdctufOmcIejCbGI9s2PMfQPuRAAmWHo4p5b886bxAYPaNp/XZrTO5/BH8iFAMgAQxf3fTqP7X3QnAG99IVz2OQZ0uGRpLM0TfzkxlyqP1rw2ctyDM73hHwIgM765lXRi3mzh+YN7qNfMoMrBDIiADrlUxflXHMuXXu05IbzIwHm/UNKBEDHXTg19Hm2cUeL8nM0pgVDWrwE7qCRA40f3RDzuhWZZDti3xG74rBTccTZf8SuOOJUHHFqGpxESiRSbiLtJpJuIi3SlhsytVBQhE0tHNRCphYOipyQ1ic/UFQYKCoIFBXqfQoCRQWBPgUB1kJfclooHJT9+f9AlbN9v73roH2gyjlQ7Rysdmobnbq4W59w40k3bQvbFpbtBgIiHNTCphYKaiFDhINaflTrkx/oXRDonR/okx/onR/o3zPQv6dOj8cvCICOKIwFfv2ZvIj0N3bLXFfsPGCv2Wmt3WGt3ZnesNtOptt0ino85cZTQohW/sd6QJQU6WXFRtkAfUSxUTZAH9hL1/z9mbXbFbNl3PKzMeku25RetS29aru1fpfVkGjT9+7Yoj7u1sdb+R+Hg9qQIr20n17aXy/tp5f2M4b21XkUkBMB0G6aJu66Obe4p1+v6JpG95U1qRfeTb22PlXX2s3cGbYj3ttnv7fPfubtY/9NJKSVFeuTh5nTysxpw818/5zZ2zGl/fRRAyW6xRoS7lPLk4tXppZtTqWtrvopiZS7cY+1cc/xH5AT0sYPMSYMMSaUmhOGGKyHkIdEV6dffHJeZOZo/x3sdajWeWZF8vl3U8s3p23HmzbEk+7q7dbq7dYDz8U1TZQVG9PKjOkjzGllZs/cLCwKF0izLLz8sPPbpxufWJZsTHZh5DfnaIdj2aa0EHEhRHHPwORh5qzR5qzRwaLCLPzefYQAaJ9xJcbnLvPZgv51O62/vBh/anmy6x76OsB1xea91ua91t9eTGiaGDPIOGdC8JwJwTGDjKwZJporwcrwxqT780ca//lKXJ5vv/ywU344+eRbSSHE8P76GWOCs0ab00aYfh9T9SMCoB0iIe2uW3L9sqLHdsQzbyf/+kJ81XvS3PrNcF2xbpe1bpd17xONRYWBC6eG5k0NThpq+joJCmOB0YM8vr9Wbk1/8Q/15Ydsb5vRgq0V9taK+ANL4qYhTh8ZvGBK8LxJwcIY3YJuQgC0w+cvyxnU2x/l//UN6R8+WL+lXN47vzkHqpw/LYn/aUm8qDBw6YzQZTPDw/v74zM/xemjPA6wJ99Kfu2B+pTlwZhPB6QtsXRdaum61Lf/KqaVmfOmhuZOCvK2oKsRAG01rsT4xHk+WPO1fb/9o383vLwm5XVDOutAlfP7Z+O/fzY+drBx+czQh04P5+X4qUcwrczLF0XPrkx+8f461x/F/yS2I46+MLjz72LSMPOy00Pzp4VyI3766n2EAGgTQxc/uC4m+ezmtCV+/mjDn5bEvXrH20WOjg799OHG+dNCV50VnjjUHxfthCGetXPDbutL99f7sfqfyHHFyq3plVvTi/7VcP7k4IJZ4dNG+ntUUEL+uJc894lzIyMGSP1Z7Txgf/53det3yT7c32GJlPvI64lHXk+MLTGunxu5cGpI5pcxpiG8umBsR3z1gfo2LunwhUTKfWJZ8ollyeKegQ/PCl8+MzSgl8Tfva8wxNa6wljgNrmPeHzszeRld1ZncfU/0bqd1h2/rzv7K0d++0y8S9cxdEZpP8P06IHh4dcSm/Zk55VQfti594nGc79WdfM9ta+tT/u9iyMDqZ9qJfFfH8qRdvTZdsQ3/1z/8OsJrxvS3Q5WO//zcMPvnm68+pzwdXMjsi0jGNrXm0dU1xX3L4578qO7jeuKl9ekXl6TGtpXv+bcyOUzQzkhSW9P+cl120hoaF/9qrNkXM0vhEik3Nt+Watg9X9fXdz97dPxc75S9YMHGw7VSvTqo7SfNwGwZoe184D/pn51zPb99vf+Xj/7i0d++O+G3QdV+a0zix5AK/7rQzlybmNSH3dvubd2xZa01w3xXiLl/mlJ/MFXEh+fE/7kvIgMs8i9GqR+cbXvZ3+1V13cfeC5LO/0dB3vbxWZDe2rz5sqy2r+Ex2pcxb+tIbqf6J4yr3/2fg5X6n65ZONcS82PDhRUYE3d9aaHVwSaAd6AC257eIcCad+xlPuTb+o3bA7O1/0dVJj0r3n8cZ/vpy4/dKctHejAl5tcePHpX/wEAHQrMF99IvkO+vRdcUX769bt5Pq35LKGufbf633sAGFMQ8eHFxXHJbpRQjkxxBQsz5xbljC0f+fPtyw5B3lxnl9Jy/Hg0snnnIdZkaiPeSrcHLICWmXz5Ru8s9Drybuf5b3XbKLBDVPFqll2QpwdAMCoGmXzwzFJNt+ZPt++85/NHjdCrTOqzMgmQ6P9iIAmrZwjlz7vtmO+Mof67JpfX8W82qPCj0g+ng0+wg+xeXShAlDDa8W8jTnD4vjq7fz4tcfTO+uHZ9unQ2vEABNuOx0uUb/t1XY9zze6HUr0Fa2d/2000Z5fwYZfIQAOJWhi/nT5LqLvvM33xzrASFEIuXZlzV/apANk9F2BMCpzhwr14l0r65PseLXXxLeTdMd2FufM0GuxxfITKJKJ4l5U+Ra/HXXowz++Ewy7eV8/Ds+HJX5pARIhQA4SUATZ4338iS/Uyx5J8WiXz/ycL7WsP76FxdEvfrp8BcC4CSThpnyjP84rvjF40z896Ujnm7JcMP5EWn3MIdUZCl2kjhXpvHT1zek2NvLpyqOeLwq93sLY7dcKNdaFkiIADjJ7LESjf88tDTpdRPQQZ4HgKaJOxZEf3N7Xu987nE0i4vjuPyoVlYsy/aoR+qc598lAPyq4rAUXbc5E4KLv19460U5EY92p4DkCIDjpo8w5ZlD/egbSUuKGoKO2HtIlo3ZYhHtC5fnvPyTwtsvzaE3gFNwQRw3Y4RE4z//flXdk36zwEbJjuspjAVuvzRn6U97/OozeXMnB0OmNE868JQsIx4ymDpclgDYtMfasZ/nfx/btNdKWW7QkKvO6gFx3sTgeRODDQn3pTWpl1anXl2fqq5nkbm6CIBjgoY2YoAsn8Yr61j662+WLTbuticMleWKOkU0rF08PXTx9JDjirU7rGWb029tSq3cZnl+ljK6maQXaPcbMUCX5/yvV9dx5pfvrd1pSRsA7wtoYsJQY8JQ45YLI7YjtpRbq96z3t2eXrfT2r7P5nyxrCf7BdptxpbI8lE0JNx3ttED8L03N6UWzvHTaiw9IEYNNEYNNK4+OyyEaEy6G3Zba3da63dZa3daOw/YLnmQdWSpep4bM0iWj+KNDWnm/2SB19alEynXq9PBOi8npE0dbr7/Yqw+7q7fba3beSwSdh3kGs0GslQ9zw2XZgXAaxsY/8kG8ZT7xsZ01uzNGYtoM0aY78+Uq210j/YM1u201u6yyg+RB74kS9Xz3NC+suyguJbd37LF86tSWRMAp8jL0U4fZZ4+6lgeVNe763ZZa3em1+yw1uywKmtkWQaBlhEAQgiRH9Xyo1J01W1HbGX/n2yxZFXyOx+PqjDpviCmnTHGPGPMsTzYd8RZ9V565TZr5db05r2WTRzIigAQQqbH/+37bU5+zxo1De6TbyU/coafXgVnRL8egX49QvOnhYQQdXF3xZb0W5vSb2xMb95L71YuBIAQQgwukiUAZFtBik766wsJBQPgRLkRbc6E4NGhsMoaZ+m61NK16VfXp+rjPOh4jwAQQoj+PaQJgD0EQFbZuMdauTU9RZpF5t7qnR9YMCu8YFY4bYllm1MvrEotXpk6XMcIkWekWfvkqaICWT4H+sjZ59dPx71ugnRMQ8weE/zuwthr/9Pjz3fkXzE7HItk/5sSCclS+LzVR5oAqDjM01C2Wbo29cYGVvY1TQ+I00eZi66NvfE/PX5+c+7poyTakVcFshQ+bxUVyvI5HKgmALLQD//dwLYKLQsHtYunh/58R/7i7xdef34klw5Bt5Cl8Hmrjxz7pDck3IYEdSILbd5rPfwa+3u3SUmR/rWPRl/9WY+vXxkt7inFjZnFeAkshBB5OVI8bkj4+H/m2OD9n8vzuhUZ8+fn44v+1eDJj/7xQw1njAn260FFa5OckHbd3MjCOZFH30j8+qn4XlYadw0uR2HoQpINWw5USRcAyJTaRvfLf6hjIKhdDF1cMTv83KLCb30s2iOXYpV5fKYiL0eWD4EAyG5vbU7f/ywzgtrN0MU150Ze+GHhDedH5NmzPTvwcQp5XjfV8wIg2939WMOyTcwI6ohoWPvqR6NPfLdwUikD1xlDAAh5JiAnUgRAlrNscdt9taz26LDh/fV/frXgqx+Nynbcpk8RAMKUZRWwSFoEQParj7s3/aJ23xGG+zoooIkbzo889I38Eml2cPEvAkAYuiyPEvQAFHGgyrn+5zW88umMUQONx75dcN6k7Nxtu9sQAMKQ5jEixeCwMrbvt6/8YfWO/cxu7LickHbfbXm3XBjxuiE+RgAIU54eABtBq6TiiHPVj2rW7OB9QMdpmrhjQfQbV0XZQKJjCACh0wOAR6rqnWt+VvOf5UmvG+Jv154X+epHo163wpcIAGFL0ws3md6mnnjS/cLv6r75l3reAHXG9XMjt8xnLKjdCACRtmW58VQ4OxBN+vfSxEd+UPPePmkeRnzojg9Hr5it9Nk7HUAACEuamy7MqSEK27LXuuzO6vv+05hiNnBH3XlNbOJQ+tHtQAAIS5oeQJAegNqSafcXjzVe/J1qVgt3jB4QP785V56lnfIjAERamh4AQ0AQQuw8YH/iZzWf/13droPSXJr+MaCXfuc1Ma9b4RsEgJBnC/4QQ0D4f08tT877ZtXX/1Rfzk7I7XTx9ND8aSGvW+EPBICoa5QlAKJhegA4znbE/76WmPuNqm//tZ4N8dvlSx+J0p9uCwJA1MZlWZEvz9HEkIdli3+9kjjva1W33lv7+oa0K8vjitSKewauO48ZQa2j4oi0JZJyLMHtSwCgGY4rXlyduv7nNRd+q+pvLyZqpem2SuuW+Tk9OUOmNXxAQgghye3Up0CaRcmQ1fb99p3/qJ/5hSOf/lXtsyuTkjy7SCgW0a6by9KwVjBnVgghKmuc3hKcC18Q04KGxjRwtCpluUveSS15JxUNa+dPDs6fFjptpMmo9yk+embol082kpEtIACEEOJgtTN6kNeNEEII0bsgwKwPtF1Dwn30jeSjbyTDQe20keZZ44JnjTMH9KIrKYQQhbHARdNDj7ye8Loh8iIAhBDiYLUs74H7FRIA6IhEyn15TerlNSkhxLD++sxRwallxpRhpgxdWw8tnBMmAFpAAAghxAFpAmB4sf72VlaBolO2VdjbKuJ/eUEIIQb30acON4+GgYJHaI0dbAzrr2+r4KGqaQSAEEJUHJYlAEYO4BtBJu06aO86aD/8uhBCFMS0SUPNSaXGxFJz/BAjJ6TEO4NzJwS3VcS9boWkKDdCCCHPmvuRA/lG0FWq692X1qReWpMSQugBUVZsTB5mTCw1J5Uag3pnbefg3Emh3z5DADSNciOEEDsPyBIAIwbomiZY7IOuZjti4x5r4x7r7y8lhBA9cgNHewaThhrjhhiRYPZ0DsYPMXrlBQ7VytLLlwoBIIQQh2qduribK8EmgjkhbWAvfXelLIEERRypc154N/XCu8c6ByMHGhOHGpOHmVOHm/16+Ps1ckATZ4wxH3uTY9eaQAAcs/OAPa5Eik9jbIlBAMBDtiPW77LW7zrWOSjuGZg63Jw+wpw5Oljc05dhMH4IAdA0KUqeDLaUW5IEwMxR5tMruFghi/LDTvnh5OPLkkKIkiJ91mhzzsTgaSOCPjrBdKwct7aE+FyOWb/LWjDL60YIIYSYNSbodROOW7ouVXbToQz+hfOmhO75VG4G/0J0p50H7J0H7L+/lIiGtTPHBi+eHjprvBk0vB87bdnIAboeEDZvAT6AADhmwy5ZRl2KewZKinR53ksDH9SQcJ95O/nM28nciDZvauhjZ4VlfsoOB7WyYmPjHsvrhkjHlyN6XWHjHsuRZu7NGWM4Ggb+UBd3H3o18eHvV39kUfXilUlpJ7AN6Zu181w7gwA4Jp5y5VkueIZMo0BAW6zZYd3+67qLvlN1dJ2BbPw+l6mL8KEcJ88eDLNGm3k5so+rAh+0rcK+5Z7am++prayRa8SdAGgSH8pxyzfLEgAhU7toOoeawq9eXpO69LvVK7bIckMJIfoVMgTUBALgOKmu1wWzONAOPna4zrnx7trX1styT9EDaBIfynGVNc6O/bK8Bhg/xBjWn2cW+Fgi5d7+69ot5VLcU3lRxlSbQACcZOk6id5f0QmA3zUk3M/9tjYtwfTLbNrdKIMIgJO8uFqiAPjIGeFomKsW/ratwv7jEu834+S8zCYRACdZsSVdH5dlJnN+VLvmXE61hu/9cXE8nvT4tgozs7opBMBJLFu8ul6iTsCN50foBMDvquqdxSs9vq2ChhbgTvoAAuBUz74tUQDkR7VP0AmA/z0lwf6GAardB/CRnOrF1Sl5RoGEEDecH4lJcFAB0Blvb017vheb5w2QEAFwqmTafe4d759W3pcf1b60IOp1K4BOaUi43s6xth0O2msCAdCEo1ufy+Oqs8LTytgeDv6255CXAZC2KP9NIACa8Nam9F5PL9ZTaJpYdG2MeWzwtYPVXg7B1CcIgCYQAE1wXPHPlxNet+IkJUX6Zy7N8boVQMd5OxNUqhd78iAAmvbQa4lkWq4r5qYLIqePYiAIfuXtEHwdAdAUAqBp1fXu0yskmg8qhNAD4p5b8wb2ZoMgf7hjQfSK2WGdO+z/eTuZ7Ugdc4CawOXZrD8+F5dt2kB+VPvN7XksDfOFfoWBRdfGnr6zcN6UkMY3JkRhrpfVxts3ENIiAJq1ea8l4dlGw/vrP7spl4LiF0P66vd8KveRbxbMmaD6XgTDPd3d9gAB0BQCoCW/fqrR6yY04dyJwUXXxljX7iNjBhu/uT3vkW+pGwO5Ec3b0ct9RwiAJhAALVm93Xp9gywnWpzoI2eEf3xjLuPL/jJ2sPGb2/Me/VbBvCkh1fL7rHFBb3/lXQclmtgtD0pIK+56tEG2NwFHfei00M9vJgP8Z8xg455P5T77/cIrZodNw+vWdJf5Xh9xKs9ZT1KhfrRizQ7raQn2sWrShVND9306Lyek2MNkVigp0hddG3vpRz1uvSinIJbl3+CQvrq3Y1/1cVe2Q+olQQC07mePNKZkXUc+Z0LwkW8VlBUzN9SX+hQEvnB5ztKf9PjvT8Sy+ATQz1+e4+34z4bdEpxJJiUCoHXlh+y/PC/XwuATDe2rP/SNgstO97iLjQ4LB7Urzww/fWfhX7+UP39ayMiuIDhvYnDeFI8vzvUEQDMIgDa598nGColnEUSC2k9uzF10bSyXjaP9bMYI8+5bcpf+tMcdH46WFGVDDgzpq//whlyvWyHW7iAAmkYAtEk86X7vb/Vet6IVV8wOL15U+OFZYVYJ+FqvvMAt8yPPLSr811fzr5jt43Oh+/cI3P/ZvPwc79u/fIuMc/lkQAC01UtrUs+ulPRt8Pt65QV+dH3swa8VjB2szPyS7DV5mLno2tiyu3rc+6ncC6aE/LUdbFmx/uDXCmTYuWTXQZtlwM2hTLTDf/+jYcYIszAme2pOHGr87zcLnn07+cCS+OrtdH79LWRqF0wJXTAl1JBwX1yden5V6pW1qUavz1hv2ZVnhr9+VTQSlCKx3tzI43+zCIB2qKxxvv6n+l9/Js/rhrQuoIn500Lzp4XeeS/9x8Xx51elHKkrBloXDWuXzAhdMiOUstxlG9Mvrk69uj69p1Ku6e1lA4yvfzQ6c7RE29a+uFq6DV3kQQC0zwvvpv71SuKqs8JeN6StJpeak28z9x6yn1iWfO6dlLfz4Ur76XMmKroXQgYFDe3MccEzxwWFEDsP2K+tT7+5KfX2Fquq3suBjlEDjevPj1wyIyTV4sR40qUH0AICoN1++GDDtDKztJ/3g5ttN6CXftvFObddnLP3kL14Zeq5d5Jrdljdc0Z2QBNjS4yzxgXPnRgcPYjrLcNKivSSIn3hnLDrim377BVb0qu3p9futLbvs7unz1cQ0+ZNCV0yIyTnqaUvrUnJdrCHVLgh2y2ecj99X+3/fqPA2/3NO2ZAL/3GCyI3XhCJJ921O63VO6zV29Ord1gHqjKZBkWFgbL+xtgSY/IwY2KpKcM8kKynaWJ4f314f/3qs8NCiMaku2G3tXmvvbXc2lJub99vZ3BD/FhEGzPImDHSPH2kObHUlOqR/xSPviH7xA1vEQAdsX2//eU/1N336Tz/TriMhLTpI8zpI0whIkKIw3VO+SGn/LBdcfjYvw/XOfGUSCTdRNpNpNxESli2GzK1oClCphYytbApQqZWENOKCvSiwkCfgkBRQaCoMDCot55HxfdaTkibOtycOvz4U3lDwt1dae+pdPZX2ZXVzsEa53CtW9Po1DW6dXE3kXItW6Rt13ZEQBOmoZm6CAe1gqiWHw0UxLTinnpxz8CA3vrIAXpxT90XV/6hWue19bwAaAkB0EHPv5u67z+Nn7kkS87p7Zkb6JkbGD+E6yFrRcPaqIHGqIFet6MbPbg00T3jnP4lcedNevc+0fjUcjqYgIwsW/zjJXl3cJEEAdBxriu+/Me6ZZuYYwBI58m3kuwA2ioCoFPSlrjtvtrNe1lsBUjEdsSv/iPjcX6yIQA6qz7u3nBXLcdNAPJ49I0ER4C1BQGQAZU1zsKf1uw8wAUHeK8x6d79GI//bUIAZAYZAEjiV/9pZPe3NiIAMuZgtfPxn9TwPgDw0JZy+4Elca9b4RsEQCZV1jhX/7hm+WbmBQEesB3xlT/WpXkGazMCIMPq4u4Nd9Uulv7kACD73PN44/pdlP92IAAyL2W5n/1N3e+eoR8KdJ9X1qZ+8zTvftuHAOgSjit+9nDD539Xl0ixEyHQ5XYdtL/4+zqXu62dCIAu9NTy5FU/qik/xNQgoAtV1Ts33V1b00j5bzcCoGtt2G196M7qZ97mlQDQJRqT7s331LLsq2MIgC5X2+h+9jd13/hzfZzhICCjGhLuDXfVcvB1hxEA3eShVxOX3Vn9zjZmiHpm3xHnLWboZpGaBveGu2q4pzqD/d+7z4799sd+XHP12eEvLohGw344UCNbvLvd+tOS+OKVSXaHzxp7D9k33s0eXJ1FAHQr1xV/fynx4urUtz4WO28Sx6N3LdsRz76dfGBJfM0OhgiyyrJN6c//tu5w5g65VBYB4IF9R5zb7qs9fZT5jSujZQP4CjKvut598NXE316MZ/asY3jOdcXvno3f/WgDnbmM0IbfWOl1G9SlB8QVs8O3X5rTO5+XMRnguOLNjelHXk8sWZXyfAXGoN76h04PXTg1NKy/7m1Lskb5IfurD9TzIieDCADvhYPa1WeHb5oX6ZVHDHTQ9v32o68nHluWlPCRf3h/ff700IVTQ0P7kgQdZDviny8nfvZwQ2OSqXSZRADIIhLUrj4nfMP5EXoDbVfT4D61IvnI6wlfjPKPGGBcNC04b2qopIgkaIeVW9Pf+0fDpj0++Ip9hwCQi2mIi6aFrj0vMmYw7waaVVXvvLwmveSd5NJ16ZTlv0fCkiL9nPHBs8cHpw43Tb7n5m3ea939WOML76a8bkjWIgAkNXW4efU54bmTgiGTCaPH7K60X1iVev7d1Dvb0tnxDjAa1maNNs8eHzxrXJCe34nW7LD+sDj+7Mok2/t0KQJAank52kXTQwtmhccPUfRB0XHFup3W8++mXliV3FqRtZO+NU2MHmTMHGXOGGFOGW4qu0zEssUL7yb/+mKCQzW6BwHgDyVF+gVTghdMCY1VYGjIccXmvdZbm9LLt6RXbE6rtsmXHhCjBxkzRpozRphThpmxiBJhsKXcfvzNxCOvJ5nd350IAJ8p7hmYOzk0e6w5bbgZDmZPaXBcsXG3tXxzevmW9Iot6VrFin5z9IAYNciYXGqOG2KMKzFKivRA9nznQgixea+1ZFXq6RXJbdnbvZMZAeBXQUObMtyYNTo4fYQ5ZpDhu3eJtiPe22dv2G1t3G1t3GOt22XVxyn6rYiGtbGDjbElxrgSY1yJMbC3L2cTHa5zVmxJv74+/cra1H75pu0qhQDIBkFDGzvYmFhqTBhqjBpoDOoj43NiVb2z44C9cbe9Ybe1cY+1Za/txwk8UolFtKF99dJ+emk/o7SfXtpPH9hb1+V7l2w7Ymu5tXantXqHtXJr+r19POzLggDIQuGgNqy/PnKAMay/Pqi3PrB3YEAvvdveK6YtUX7Y3nPI3lPp7Km091Taeyqd3ZV2Q4Jy3+WChlZSFBjS1+jfI9Dv2D96vx6BXnkBrbueCRoS7t5D9p5K57399rZya9s+e1uFnUzz7cuIAFBFz9xAvx6B3vmBXvmBXnmB3vmBgpiWG9FiYS0W0WJhLRLSDF0zdKEHhKFrAU3YjkhZbtoSKctNpd20LVJpN2WLlOXWx92qevdInVNV71bVOVX1zvv/sbrecbjZJWMaoqhA71sYKMzVCqKBvBytIBrIj2pH/5CXo4WDmmkI09BMXQQNzdCFaQhD1zQhHFfYjms7Im2JRMqNp9xEym1IuDWNbk2DW9PgVNW7lTXOwWqnssbZd8SpqmdUxzcIAABQlHzjhQCAbkEAAICiCAAAUBQBAACKIgAAQFEEAAAoigAAAEURAACgKAIAABRFAACAoggAAFAUAQAAiiIAAEBRBAAAKIoAAABFEQAAoCgCAAAURQAAgKIIAABQFAEAAIoiAABAUQQAACiKAAAARREAAKAoAgAAFEUAAICiCAAAUBQBAACKIgAAQFEEAAAoigAAAEURAACgKAIAABRFAACAoggAAFAUAQAAiiIAAEBRBAAAKIoAAABFEQAAoCgCAAAURQAAgKIIAABQFAEAAIoiAABAUQQAACiKAAAARREAAKAoAgAAFEUAAICiCAAAUBQBAACKIgAAQFEEAAAoigAAAEURAACgKAIAABRFAACAoggAAFAUAQAAiiIAAEBRBAAAKIoAAABFEQAAoCjD6wYga225v1dn/u9lNx3acn+vspsOZao9AE6hDb+x0us2wPc6Wevbi1QAMoIAQLt1c7lvCyIB6AACAG0iYdFvDmEAtBEBgGb5qOg3hzAAWkAA4LgsqPgtIw+AExEAqsv6ot8cwgAgAFSkbNFvgbJ5wFxblbEOABDi5FDM4oJI9uNEBABwqiwLA4o+mkMAAC05pXr6Ig+o+Ggj3gEoihrReVKFQec33shUS+Aj9ACADvJ8pIgURycRAEAGNFmLM7Kf3dG/gVqPrkAAoKtk5KE4CwpfFvwKyFYEADqlq4c+Wvj7KaxAJ/ESWFGdqZ4SvjAkDDpJwu8U3YAAUFe7iqaPCgRh0DE++oqRKQwBoWn+LQentJw8AJpDAOA4/xb9Fpz4SxEGwIkIAKVlZcVvAZ0D4ES8AwCEUDIMVIt/fBABAJwqi8OAoo8TEQBAK3ydB1R8tIAAANrBF2FA0UcbEQBAB0kVBhR9dAABAGRGN+cBFR+dRwAAXYud+iEtAgAAFBXwugEAAG8QAACgKAIAABRFAACAoggAAFAUAQAAiiIAAEBRBAAAKIoAAABFcSIYukmHd0Qou+nQlvt7sSMCkHFsBYEu1EX7oxEGQEYQAMgASTZGJhiAdiEA0EGSFP3mEAZAqwgAtIPkRb85hAHQJAIArfBp0W8OYQC8jwBAE7Ks6DeHMIDiCAAco0jRb052hwHzaNEk1gEAQpyQf1lTKBVPdLQFK4GBk2RH3cyO3wJdjSEgHEfVOIXvegMtf4O++3XQ1RgCApp1Yj2VtnoS2+gwAgBokw/W2W7epOjoz6LcI4MIAKBTurSXQLlHlyIAcFyXPmB2sjj6ohRK3kgmg+IUvATGSTJbwrqo3EheZ2VGAOBE9ACQAd1cVpr7cQQD0C4EADpIwmfJE5tEGACtYggIp2qhdEpY9NuCMDiRT79EdAV6AGhFFtQLegZAkwgANCELin5z3v/VSAKAISCoToUkyOJER2cQAMAxWZYE3bxQGX5EAAAnyZoYoPSjVQQA0CzfhQFFH+1CAABtIm0YUPTRYQQA0CndFgwUemQcAQBkTMbDgKKPLkUAAICiOBMYABRFAACAoggAAFAUAQAAiiIAAEBRBAAAKIoAAABFEQAAoCgCAAAURQAAgKI4EhLdJ7Nb5bBPDtBJ7AWErtU9m2USBkAHEADImKMHEMqzbz6pALSMAEBnyVPxW0AYAB9EAKAjfFH0m0QSAO9jFhDazb/VX/i88UBm0QNAW2Vf6aQ3AMURAGhF9tX9JmVHGBx9D+91K+AbrAMAhDgh5/xYQBUJaWQcPQC0Ttn6InMYtPClyNxsSIUeANAs2boFyiYxuggBALTuxMrb/WFA3UcXYQgIbUINas7Rxc+ZSoVMraaWpMsCydEDADKgM10EwhVeIQDQtU55ns3gk6m0dVPahgGnIADQHbpiROLEv9PDmpvxH01+oNvwDgBt1cbCJMnoc3OtlWq/0i4lyRcBmdEDQGfJWWjeb9Up5V6R6g+0BQGAjpOz9J/ixCRQ5/EfaAt2A0VbnVLufVH9T+S7BncSUYdW0QNA+/i9jDY3NAQoiJfAUF32JYHfQxrdhgAAjvNpGFDx0TEEANAEXyQBdR+dRAAArZAtDKj7yBQCAGgrryaSUvHRRQgAoLM6GQldt10S0DKmgQKdlZG1ZtR9dD96AACgKFYCA4CiCAAAUBQBAACKIgAAQFEEAAAoigAAAEURAACgKAIAABRFAACAoggAAFAUAQAAiiIAAEBRBAAAKIoAAABFEQAAoCgCAAAURQAAgKIIAABQFAEAAIoiAABAUQQAACiKAAAARREAAKAoAgAAFEUAAICiCAAAUBQBAACKIgAAQFEEAAAoigAAAEURAACgKAIAABRFAACAoggAAFAUAQAAiiIAAEBRBAAAKIoAAABFEQAAoCgCAAAURQAAgKIIAABQFAEAAIoiAABAUQQAACiKAAAARREAAKAoAgAAFEUAAICiCAAAUBQBAACKIgAAQFEEAAAoigAAAEURAACgKAIAABRFAACAoggAAFAUAQAAiiIAAEBRBAAAKOr/AMy3e06nqrxFAAAAAElFTkSuQmCC",
  "apple-touch-icon.png": "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"
};

function pwaResource(name) {
  var b64 = PWA_ICON_BASE64[name];
  if (!b64) return null;
  return {
    b64: b64,
    mime: "image/png"
  };
}

function pwaManifest() {
  var manifest = {
    name: "Träningsnärvaro Göteborg Sim",
    short_name: "Närvaro GS",
    description: "Mobil träningsnärvaro för Göteborg Sim kopplad till Google Sheet.",
    lang: "sv",
    display: "standalone",
    start_url: "./",
    theme_color: "#2563eb",
    background_color: "#f1f5f9",
    icons: [
      { src: "./icon-192.png", sizes: "192x192", type: "image/png", purpose: "any" },
      { src: "./icon-512.png", sizes: "512x512", type: "image/png", purpose: "any" },
      { src: "./icon-512.png", sizes: "512x512", type: "image/png", purpose: "maskable" }
    ]
  };
  return JSON.stringify(manifest);
}

function doGet(e) {

  // PWA resource requests (icons + manifest) served from the same
  // web-app URL, so the browser sees a same-domain absolute src.
  var params = (e && e.parameter) || {};
  if (params.pwa) {
    if (params.pwa === "manifest") {
      return ContentService
        .createTextOutput(pwaManifest())
        .setMimeType(ContentService.MimeType.JSON);
    }
    var res = pwaResource(params.pwa);
    if (res) {
      return ContentService
        .createTextOutput(Utilities.newBlob(
          Utilities.base64Decode(res.b64),
          res.mime,
          params.pwa
        ));
    }
    // fall through to HTML when resource name is unknown
  }

  return HtmlService
    .createHtmlOutputFromFile(
      "Index"
    )
    .setTitle(
      "Träningsnärvaro Göteborg Sim"
    )
    .addMetaTag(
      "viewport",
      "width=device-width, initial-scale=1.0"
    );
}


// ============================================================
// D — LEGACY JSON POST API
// ============================================================
//
// Kept temporarily so we do not unnecessarily remove working
// backend functionality while Option B is being validated.
//
// Option B itself does NOT use this function.
// ============================================================

function doPost(e) {

  try {

    if (
      !e ||
      !e.postData ||
      !e.postData.contents
    ) {

      return jsonResponse({

        success:
          false,

        error:
          "No POST data received"
      });
    }

    const request =
      JSON.parse(
        e.postData.contents
      );

    const action =
      request.action || "";

    if (
      action ===
      "saveAttendance"
    ) {

      if (!request.date) {

        return jsonResponse({

          success:
            false,

          error:
            "Missing date"
        });
      }

      if (!request.location) {

        return jsonResponse({

          success:
            false,

          error:
            "Missing location"
        });
      }

      if (
        !Array.isArray(
          request.attendance
        )
      ) {

        return jsonResponse({

          success:
            false,

          error:
            "attendance must be an array"
        });
      }

      const result =
        saveAttendance(
          request.date,
          request.location,
          request.attendance
        );

      return jsonResponse({

        success:
          true,

        action:
          "saveAttendance",

        data:
          result
      });
    }

    return jsonResponse({

      success:
        false,

      error:
        "Unknown action",

      action:
        action
    });

  } catch (error) {

    return jsonResponse({

      success:
        false,

      error:
        error.message
    });
  }
}


// ------------------------------------------------------------
// Legacy JSON response helper
// ------------------------------------------------------------

function jsonResponse(
  payload
) {

  return ContentService
    .createTextOutput(
      JSON.stringify(payload)
    )
    .setMimeType(
      ContentService.MimeType.JSON
    );
}


// ============================================================
// E — LEGACY API TESTS
// ============================================================
//
// testApiGet() is intentionally NOT used anymore because
// doGet() now returns HtmlService instead of JSON.
//
// The old REST GET API was already tested successfully before
// Option B was introduced.
// ============================================================


// ============================================================
// E2 — LEGACY POST API TEST
// ============================================================
//
// This test remains available for the old POST API.
//
// IMPORTANT:
// It dynamically finds Askim 4/9 rather than assuming column J.
// ============================================================

function testApiPost() {

  Logger.log(
    "=== E2 — POST API TEST ==="
  );

  const requestData = {

    action:
      "saveAttendance",

    date:
      "4/9",

    location:
      "Askim",

    attendance: [

      {
        row:
          2,

        present:
          true
      },

      {
        row:
          3,

        present:
          false
      }
    ]
  };

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(
      "HT -26"
    );

  // Find the actual Askim 4/9 column.
  const training =
    ensureTrainingColumn(
      "4/9",
      "Askim"
    );

  const targetColumn =
    training.column;

  const originalChrister =
    sheet
      .getRange(
        2,
        targetColumn
      )
      .getValue();

  const originalAmanda =
    sheet
      .getRange(
        3,
        targetColumn
      )
      .getValue();

  Logger.log(
    "Testing column: " +
    training.columnLetter
  );

  // ----------------------------------------------------------
  // Simulated POST event
  // ----------------------------------------------------------

  const mockEvent = {

    postData: {

      contents:
        JSON.stringify(
          requestData
        )
    }
  };

  const response =
    doPost(
      mockEvent
    );

  const text =
    response.getContent();

  const parsed =
    JSON.parse(text);

  if (
    !parsed.success
  ) {

    throw new Error(
      "POST API failed: " +
      parsed.error
    );
  }

  if (
    parsed.action !==
    "saveAttendance"
  ) {

    throw new Error(
      "Unexpected action returned"
    );
  }

  // ----------------------------------------------------------
  // Verify
  // ----------------------------------------------------------

  const christerAfter =
    sheet
      .getRange(
        2,
        targetColumn
      )
      .getValue();

  const amandaAfter =
    sheet
      .getRange(
        3,
        targetColumn
      )
      .getValue();

  if (
    String(christerAfter).trim()
    !==
    "x"
  ) {

    throw new Error(
      "Christer was not saved as x"
    );
  }

  if (
    String(amandaAfter).trim()
    !==
    ""
  ) {

    throw new Error(
      "Amanda was not saved as blank"
    );
  }

  // ----------------------------------------------------------
  // Restore
  // ----------------------------------------------------------

  sheet
    .getRange(
      2,
      targetColumn
    )
    .setValue(
      originalChrister
    );

  sheet
    .getRange(
      3,
      targetColumn
    )
    .setValue(
      originalAmanda
    );

  SpreadsheetApp.flush();

  Logger.log(
    "=== E2 POST TEST PASSED — VALUES RESTORED ==="
  );
}


// ============================================================
// F0 — DELETE TRAINING COLUMN
// ============================================================

function deleteTrainingColumn(
  date,
  location
) {

  const lock =
    LockService.getScriptLock();

  try {

    // --------------------------------------------------------
    // Försök erhålla lås för att förhindra samtidiga
    // strukturändringar i arket
    // --------------------------------------------------------

    const hasLock =
      lock.tryLock(
        15000
      );

    if (!hasLock) {

      return {

        success:
          false,

        deleted:
          false,

        error:
          "Kunde inte erhålla scriptlås. En annan operation pågår. Försök igen om en stund."
      };
    }

    // --------------------------------------------------------
    // 1. Validera lokal
    // --------------------------------------------------------

    if (
      !location ||
      (
        location !== "Askim" &&
        location !== "Valhalla"
      )
    ) {

      return {

        success:
          false,

        deleted:
          false,

        error:
          "Ogiltig plats: '" +
          location +
          "'. Endast 'Askim' eller 'Valhalla' är tillåtna."
      };
    }

    // --------------------------------------------------------
    // 2. Validera och normalisera datum
    // --------------------------------------------------------

    if (
      !date ||
      typeof date !== "string"
    ) {

      return {

        success:
          false,

        deleted:
          false,

        error:
          "Ogiltigt eller saknat datum."
      };
    }

    const dateMatch =
      date.match(
        /^(\d{1,2})\/(\d{1,2})$/
      );

    if (!dateMatch) {

      return {

        success:
          false,

        deleted:
          false,

        error:
          "Ogiltigt datumformat. Använd d/M eller dd/MM, exempelvis 4/9."
      };
    }

    const day =
      Number(
        dateMatch[1]
      );

    const month =
      Number(
        dateMatch[2]
      );

    if (
      day < 1 ||
      day > 31 ||
      month < 1 ||
      month > 12
    ) {

      return {

        success:
          false,

        deleted:
          false,

        error:
          "Ogiltigt datum."
      };
    }

    const normalizedDate =
      day +
      "/" +
      month;

    // --------------------------------------------------------
    // 3. Konstruera exakt förväntad rubrik
    // --------------------------------------------------------

    const expectedHeader =
      location +
      " " +
      normalizedDate;

    // --------------------------------------------------------
    // 4. Kontrollera rubrikformat
    // --------------------------------------------------------

    const trainingHeaderPattern =
      /^(Askim|Valhalla)\s+\d{1,2}\/\d{1,2}$/;

    if (
      !trainingHeaderPattern.test(
        expectedHeader
      )
    ) {

      return {

        success:
          false,

        deleted:
          false,

        date:
          normalizedDate,

        location:
          location,

        header:
          expectedHeader,

        error:
          "Rubriken '" +
          expectedHeader +
          "' matchar inte tillåtet format för träningskolumner."
      };
    }

    // --------------------------------------------------------
    // 5. Hämta rätt kalkylblad
    // --------------------------------------------------------

    const ss =
      SpreadsheetApp
        .getActiveSpreadsheet();

    if (!ss) {

      return {

        success:
          false,

        deleted:
          false,

        error:
          "Kunde inte öppna det aktiva kalkylarket."
      };
    }

    const sheet =
      resolveAttendanceSheet(ss);

    if (!sheet) {

      return {

        success:
          false,

        deleted:
          false,

        error:
          'Kunde inte hitta bladet "HT -26" i kalkylarket.'
      };
    }

    // --------------------------------------------------------
    // 6. Läs rad 1 och samla ALLA exakta matchningar
    // --------------------------------------------------------

    const lastColumn =
      sheet.getLastColumn();

    if (lastColumn < 1) {

      return {

        success:
          false,

        deleted:
          false,

        error:
          "Kalkylarket är tomt."
      };
    }

    const headerRowValues =
      sheet
        .getRange(
          1,
          1,
          1,
          lastColumn
        )
        .getValues()[0];

    const matchedColumnIndices = [];

    for (
      let colIdx = 0;
      colIdx < headerRowValues.length;
      colIdx++
    ) {

      const headerVal =
        String(
          headerRowValues[colIdx]
        ).trim();

      if (
        headerVal ===
        expectedHeader
      ) {

        matchedColumnIndices.push(
          colIdx + 1
        );
      }
    }

    // --------------------------------------------------------
    // 7. Säkerhetsregel för antal matchningar
    // --------------------------------------------------------
    //
    // 0  = inget tas bort
    // 1  = fortsätt
    // >1 = avbryt helt
    // --------------------------------------------------------

    if (
      matchedColumnIndices.length === 0
    ) {

      return {

        success:
          false,

        deleted:
          false,

        date:
          normalizedDate,

        location:
          location,

        header:
          expectedHeader,

        error:
          "Hittade ingen kolumn med rubriken '" +
          expectedHeader +
          "'. Ingen ändring gjordes."
      };
    }

    if (
      matchedColumnIndices.length > 1
    ) {

      return {

        success:
          false,

        deleted:
          false,

        date:
          normalizedDate,

        location:
          location,

        header:
          expectedHeader,

        error:
          "Hittade " +
          matchedColumnIndices.length +
          " identiska kolumner med rubriken '" +
          expectedHeader +
          "'. Radering avbröts för att skydda dataintegriteten."
      };
    }

    // --------------------------------------------------------
    // Exakt en matchning
    // --------------------------------------------------------

    const foundColumn =
      matchedColumnIndices[0];

    // --------------------------------------------------------
    // 8. Explicit skydd för systemkolumner A och B
    // --------------------------------------------------------

    if (
      foundColumn <= 2
    ) {

      return {

        success:
          false,

        deleted:
          false,

        date:
          normalizedDate,

        location:
          location,

        header:
          expectedHeader,

        error:
          "Säkerhetsspärr: Kolumn " +
          foundColumn +
          " är en systemkolumn (A eller B) och får inte raderas."
      };
    }

    // --------------------------------------------------------
    // 9. Återverifiera faktisk rubrik före deletion
    // --------------------------------------------------------

    const actualHeaderBeforeDelete =
      String(
        sheet
          .getRange(
            1,
            foundColumn
          )
          .getValue()
      ).trim();

    if (
      actualHeaderBeforeDelete !==
      expectedHeader
    ) {

      return {

        success:
          false,

        deleted:
          false,

        date:
          normalizedDate,

        location:
          location,

        header:
          expectedHeader,

        error:
          "Avbröt: Kolumn " +
          foundColumn +
          " har rubriken '" +
          actualHeaderBeforeDelete +
          "', vilket inte exakt matchar förväntad rubrik '" +
          expectedHeader +
          "'."
      };
    }

    // --------------------------------------------------------
    // 10. Kontrollera faktisk rubrik mot träningskolumn-regex
    // --------------------------------------------------------

    if (
      !trainingHeaderPattern.test(
        actualHeaderBeforeDelete
      )
    ) {

      return {

        success:
          false,

        deleted:
          false,

        date:
          normalizedDate,

        location:
          location,

        header:
          actualHeaderBeforeDelete,

        error:
          "Avbröt: Rubriken '" +
          actualHeaderBeforeDelete +
          "' är inte en giltig träningskolumn."
      };
    }

    // --------------------------------------------------------
    // 11. Spara information före borttagning
    // --------------------------------------------------------

    const beforeHeaders =
      headerRowValues.map(
        function(header) {

          return String(
            header
          ).trim();
        }
      );

    // --------------------------------------------------------
    // 12. Ta bort HELA kolumnen fysiskt
    // --------------------------------------------------------

    sheet.deleteColumn(
      foundColumn
    );

    SpreadsheetApp.flush();

    // --------------------------------------------------------
    // 13. Verifiera borttagning
    // --------------------------------------------------------

    const newLastColumn =
      sheet.getLastColumn();

    let deletedSuccessfully =
      true;

    if (
      newLastColumn > 0
    ) {

      const afterHeaders =
        sheet
          .getRange(
            1,
            1,
            1,
            newLastColumn
          )
          .getValues()[0]
          .map(
            function(header) {

              return String(
                header
              ).trim();
            }
          );

      if (
        afterHeaders.indexOf(
          expectedHeader
        ) !== -1
      ) {

        deletedSuccessfully =
          false;
      }
    }

    if (
      !deletedSuccessfully
    ) {

      throw new Error(
        "Kolumnen kunde inte verifieras som borttagen."
      );
    }

    // --------------------------------------------------------
    // 14. Returnera resultat
    // --------------------------------------------------------

    return {

      success:
        true,

      deleted:
        true,

      date:
        normalizedDate,

      location:
        location,

      header:
        expectedHeader,

      deletedColumn:
        foundColumn,

      previousColumnCount:
        beforeHeaders.length,

      newColumnCount:
        newLastColumn,

      message:
        expectedHeader +
        " har tagits bort."
    };

  } catch (err) {

    return {

      success:
        false,

      deleted:
        false,

      error:
        "Fel vid radering av träningskolumn: " +
        (err.message || String(err))
    };

  } finally {

    try {
      lock.releaseLock();
    } catch (e) {
      // Tyst ignorering om låset redan släppts
    }
  }
}


// ============================================================
// F0 TEST — OLD DELETE TEST
// ============================================================
//
// DO NOT RUN THIS TEST AGAIN.
//
// "Valhalla 4/9" was already deleted successfully.
// Running this test now would stop with an intentional error.
// ============================================================

function testDeleteTrainingColumn() {

  Logger.log(
    "=== F0 — DELETE TRAINING COLUMN TEST ==="
  );

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(
      "HT -26"
    );

  if (!sheet) {

    throw new Error(
      'Fliken "HT -26" hittades inte.'
    );
  }

  const beforeLastColumn =
    sheet.getLastColumn();

  const beforeHeaders =
    sheet
      .getRange(
        1,
        1,
        1,
        beforeLastColumn
      )
      .getValues()[0]
      .map(
        function(header) {

          return String(
            header
          ).trim();
        }
      );

  if (
    beforeHeaders.indexOf(
      "Valhalla 4/9"
    ) === -1
  ) {

    throw new Error(
      'Testet avbryts: "Valhalla 4/9" finns inte längre. Det är redan borttaget.'
    );
  }

  const result =
    deleteTrainingColumn(
      "4/9",
      "Valhalla"
    );

  if (
    !result.success ||
    !result.deleted
  ) {

    throw new Error(
      "F0 deleteTrainingColumn rapporterade inte lyckad borttagning."
    );
  }

  Logger.log(
    "=== F0 TEST PASSED ==="
  );
}