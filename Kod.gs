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
//
// Payment status is READ from column A background colors:
//   #6aa84f (green) -> "paid",  #990000 (dark red) -> "owes",
//   anything else   -> "none"
// The web app never writes to the sheet; colors are administered
// directly in Google Sheets.
// The sheet itself is resolved robustly (see resolveAttendanceSheet),
// so a renamed sheet (e.g. "VT -27") or a new spreadsheet works
// without code changes.
// ============================================================

// ------------------------------------------------------------
// P — SHEET RESOLUTION & PAYMENT STATUS (read-only)
// ------------------------------------------------------------

var PAYMENT_COLOR_PAID = "#6aa84f";
var PAYMENT_COLOR_OWES = "#990000";

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
 * Read-only: the web app never modifies the sheet.
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
// ============================================================

function doGet(e) {

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