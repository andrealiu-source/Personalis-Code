// ============================================================
// AUDIT SYNC SCRIPT
// ============================================================
// Main tab: "xM NeXT Orders"
// OHID in column B
// Order Status in column A (Active / Resolved)
// Column C = Baseline or Subsequent
// Checkbox columns: E=xM CRC, F=NYS, G=Multiple Baselines, H=Alerts Channel
// Column I = Manual checkbox (never touched by script)
// Column N = Manual/Automatic flag (never touch Manual rows)
// ============================================================


// ------------------------------------------------------------
// HELPER: Build a map of existing OHIDs in main tab
// ------------------------------------------------------------
function buildOhidMap(data, ohidCol) {
  const map = {};
  for (let i = 1; i < data.length; i++) {
    const ohid = String(data[i][ohidCol - 1]).trim();
    if (ohid) map[ohid] = i;
  }
  return map;
}

// ------------------------------------------------------------
// HELPER: Checked checkbox (TRUE) = resolved = skip
// ------------------------------------------------------------
function isCheckedResolved(value) {
  return value === true;
}

// ------------------------------------------------------------
// HELPER: "resolved" or "dupe row" dropdown = skip
// ------------------------------------------------------------
function isDropdownResolved(value) {
  const val = String(value).trim().toLowerCase();
  return val === "resolved" || val === "dupe row";
}

// ------------------------------------------------------------
// HELPER: Multiple Baselines specific skip values
// ------------------------------------------------------------
function isMultipleBaselinesResolved(value) {
  const val = String(value).trim().toLowerCase();
  return val === "resolved" || val === "dupe row";
}

// ------------------------------------------------------------
// HELPER: "resolved" or "valid second baseline" dropdown = uncheck
// ------------------------------------------------------------
function isDropdownResolvedOrValid(value) {
  const val = String(value).trim().toLowerCase();
  return val === "resolved" || val === "dupe row" || val === "valid second baseline";
}

// ------------------------------------------------------------
// HELPER: Check if a Note Type value means Manual
// ------------------------------------------------------------
function isManualValue(value) {
  return String(value).trim().toLowerCase() === "manual";
}

// ------------------------------------------------------------
// HELPER: Check if a type value means Baseline
// ------------------------------------------------------------
function isBaseline(value) {
  const val = String(value).trim().toLowerCase();
  return val === "baseline" ||
         val === "baseline - single order" ||
         val === "baseline - first in a series";
}

// ------------------------------------------------------------
// HELPER: Check if a type value means Subsequent
// ------------------------------------------------------------
function isSubsequent(value) {
  const val = String(value).trim().toLowerCase();
  return val === "subsequent" ||
         val === "subsequent - part of series";
}

// ------------------------------------------------------------
// HELPER: Evaluate Active/Resolved status based on checkboxes
// F=col6, G=col7, H=col8, I=col9, J=col10 (1-based)
// If ANY of F-J are checked -> Active
// If ALL of F-J are unchecked -> Resolved
// Only update if column B (OHID) is not blank
// Only update if column A is "Active" or "Resolved"
// ------------------------------------------------------------
function evaluateStatus(row) {
  const ohid = String(row[1]).trim(); // Column B
  if (!ohid) return null; // skip rows with no OHID

  const statusVal = String(row[0]).trim().toLowerCase();
  // Only manage Active and Resolved — leave Cancelled and any other values alone
  if (statusVal !== "active" && statusVal !== "resolved") return null;

  const f = row[4];  // Column E
  const g = row[5];  // Column F
  const h = row[6];  // Column G
  const i = row[7];  // Column H
  const j = row[8];  // Column I

  const anyChecked = (f === true) || (g === true) || (h === true) || (i === true) || (j === true);
  return anyChecked ? "Active" : "Resolved";
}

// ------------------------------------------------------------
// HELPER: Normalize type to "Baseline" or "Subsequent"
// ------------------------------------------------------------
function normalizeType(value) {
  if (isBaseline(value)) return "Baseline";
  if (isSubsequent(value)) return "Subsequent";
  return "";
}

// ------------------------------------------------------------
// CORE: Sync audit records to xM NeXT Orders
// Instead of notes, marks/unmarks a checkbox in a specific column
// Also sets column D (Baseline/Subsequent) and column A (Active/Resolved)
// checkboxCol: the column (1-based) to check/uncheck
// shouldCheck: true = check the box, false = uncheck the box
// typeValue: "Baseline" or "Subsequent" for column D (null to skip)
// ------------------------------------------------------------
function syncCheckboxToMainTab(mainSheet, records, checkboxCol, ohidCol, noteTypeCol, typeColD) {
  if (records.length === 0) return;

  const NUM_COLS = 14; // A through O
  const lastRow = Math.max(mainSheet.getLastRow(), 1);

  // Read all existing data
  const rawData = mainSheet.getRange(1, 1, lastRow, NUM_COLS).getValues();
  const data = rawData.map(row => {
    const padded = row.slice();
    while (padded.length < NUM_COLS) padded.push("");
    return padded;
  });

  const ohidMap = buildOhidMap(data, ohidCol);
  const newRows = [];

  records.forEach(({ ohid, shouldCheck, typeValue }) => {
    if (ohidMap[ohid] !== undefined) {
      // OHID exists
      const rowIndex = ohidMap[ohid];
      if (rowIndex < data.length) {
        // Respect Manual flag in column O
        if (isManualValue(data[rowIndex][noteTypeCol - 1])) return;

        // Mark or unmark the checkbox
        data[rowIndex][checkboxCol - 1] = shouldCheck;

        // Set column D if not already set
        if (typeValue && !data[rowIndex][typeColD - 1]) {
          data[rowIndex][typeColD - 1] = typeValue;
        }

        // Evaluate and update status in column A
        const newStatus = evaluateStatus(data[rowIndex]);
        if (newStatus) data[rowIndex][0] = newStatus;
      }
    } else if (shouldCheck) {
      // OHID is new and box should be checked — add new row
      const newRow = new Array(NUM_COLS).fill("");
      newRow[ohidCol - 1] = ohid;
      newRow[checkboxCol - 1] = true;
      newRow[noteTypeCol - 1] = "Automatic";
      if (typeValue) newRow[typeColD - 1] = typeValue;
      // New row: evaluate status — checkbox is checked so Active
      newRow[0] = "Active";
      newRows.push(newRow);
      ohidMap[ohid] = data.length + newRows.length - 1;
    }
    // If shouldCheck is false and OHID doesn't exist, nothing to do
  });

  // Write updated rows back
  mainSheet.getRange(1, 1, data.length, NUM_COLS).setValues(data);

  // Append new rows if any
  if (newRows.length > 0) {
    const lastDataRow = mainSheet.getLastRow();
    let startRow = lastDataRow + 1;
    const ohidColData = mainSheet.getRange(2, ohidCol, lastDataRow, 1).getValues();
    for (let i = 0; i < ohidColData.length; i++) {
      if (String(ohidColData[i][0]).trim() === "") {
        startRow = i + 2;
        break;
      }
    }
    mainSheet.getRange(startRow, 1, newRows.length, NUM_COLS).setValues(newRows);
  }
}


// ============================================================
// ONEDIT TRIGGER: Auto-update column A status when checkboxes
// in columns F-J of "xM NeXT Orders" are manually changed
// Install this as an installable trigger (not just onEdit)
// ============================================================
function onEditTrigger(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  if (sheet.getName() !== "xM NeXT Orders") return;

  const col = e.range.getColumn();
  const row = e.range.getRow();
  const NUM_COLS = 14;

  // ----------------------------------------------------------
  // NEW: Column B (OHID) edited -> set A="Active", N="Automatic"
  // Loops over the whole edited range so bulk PASTES of OHIDs
  // (which fire onEdit once for the entire pasted block) all get
  // handled, not just the top row.
  // Runs and returns BEFORE the E-I gate below, so checkbox edits
  // in columns E-I are completely unaffected.
  // ----------------------------------------------------------
  const numRows = e.range.getNumRows();
  const numCols = e.range.getNumColumns();
  if (col <= 2 && col + numCols - 1 >= 2) { // edited range includes column B
    for (let r = row; r < row + numRows; r++) {
      if (r < 2) continue; // skip header
      const rd = sheet.getRange(r, 1, 1, NUM_COLS).getValues()[0];
      const ohid = String(rd[1]).trim(); // Column B
      if (!ohid) continue; // blank B -> leave row alone

      // Column A: Active, unless already Cancelled
      if (String(rd[0]).trim().toLowerCase() !== "cancelled") {
        sheet.getRange(r, 1).setValue("Active");
      }
      // Column N: Automatic, only if blank (don't clobber "Manual")
      if (String(rd[13]).trim() === "") {
        sheet.getRange(r, 14).setValue("Automatic");
      }
    }
    return; // done — never falls through to the E-I logic
  }

  // Only react to edits in columns E(5) through I(9)
  if (col < 5 || col > 9) return;
  // Skip header row
  if (row < 2) return;

  const rowData = sheet.getRange(row, 1, 1, NUM_COLS).getValues()[0];

  // Only update if OHID exists in column B
  const ohid = String(rowData[1]).trim();
  if (!ohid) return;

  // Respect Manual flag in column N
  if (isManualValue(rowData[13])) return;

  const newStatus = evaluateStatus(rowData);
  if (newStatus) {
    sheet.getRange(row, 1).setValue(newStatus);
  }
}


// ============================================================
// BUTTON FUNCTION: Multiple Baselines Audit
// Feeds into xM NeXT Orders only.
// OHIDs from col E. Checkbox col H.
// Check if col K = "resolved" or "valid second baseline" -> uncheck
// Otherwise -> check
// All Multiple Baselines = Baseline
// ============================================================
function syncMultipleBaselinesAudit() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const auditSheet = ss.getSheetByName("Multiple Baselines Audit");
  const mainSheet = ss.getSheetByName("xM NeXT Orders");

  if (!auditSheet || !mainSheet) {
    SpreadsheetApp.getUi().alert("Could not find required sheets. Please check sheet names.");
    return;
  }

  const data = auditSheet.getDataRange().getValues();
  const records = [];

  for (let i = 1; i < data.length; i++) {
    const status = data[i][14]; // Column O dropdown
    const shouldCheck = !isMultipleBaselinesResolved(status);

    const ohidG = String(data[i][6]).trim(); // Column G
    if (!ohidG) continue;

    records.push({
      ohid: ohidG,
      shouldCheck: shouldCheck,
      typeValue: "Baseline"
    });
  }

  // xM NeXT Orders: OHID=B(2), Checkbox=G(7), NoteType=N(14), TypeD=C(3)
  syncCheckboxToMainTab(mainSheet, records, 7, 2, 14, 3);
  SpreadsheetApp.getUi().alert("Multiple Baselines Audit sync complete. " + records.length + " active records processed.");
}


// ============================================================
// BUTTON FUNCTION: xM CRC vs Next Audit
// OHIDs from col A. Checkbox col F.
// Col B determines Baseline/Subsequent for col D.
// Skip entirely if col I checkbox is checked -> uncheck in main
// ============================================================
function syncXmCrcAudit() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const auditSheet = ss.getSheetByName("xM CRC vs Next Audit");
  const mainSheet = ss.getSheetByName("xM NeXT Orders");

  if (!auditSheet || !mainSheet) {
    SpreadsheetApp.getUi().alert("Could not find required sheets. Please check sheet names.");
    return;
  }

  const data = auditSheet.getDataRange().getValues();
  const records = [];

  for (let i = 1; i < data.length; i++) {
    const ohid = String(data[i][0]).trim(); // Column A
    if (!ohid) continue;

    const dupeStatus = String(data[i][6]).trim().toLowerCase(); // Column G dropdown
    if (dupeStatus === "dupe row") continue; // skip dupe rows entirely

    const resolved = data[i][8]; // Column I checkbox
    const type = String(data[i][1]).trim(); // Column B
    const typeValue = normalizeType(type);
    const shouldCheck = !isCheckedResolved(resolved);

    records.push({ ohid, shouldCheck, typeValue });
  }

  // xM NeXT Orders: OHID=B(2), Checkbox=E(5), NoteType=N(14), TypeD=C(3)
  syncCheckboxToMainTab(mainSheet, records, 5, 2, 14, 3);
  SpreadsheetApp.getUi().alert("xM CRC vs Next Audit sync complete. " + records.length + " records processed.");
}


// ============================================================
// BUTTON FUNCTION: NYS Audit
// OHIDs from col A. Checkbox col G.
// Col B determines Baseline/Subsequent for col D.
// Data starts on row 3 (2 header rows).
// Uncheck if col M checkbox is checked.
// ============================================================
function syncNysAudit() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const auditSheet = ss.getSheetByName("NYS Audit");
  const mainSheet = ss.getSheetByName("xM NeXT Orders");

  if (!auditSheet || !mainSheet) {
    SpreadsheetApp.getUi().alert("Could not find required sheets. Please check sheet names.");
    return;
  }

  const data = auditSheet.getDataRange().getValues();
  const records = [];

  for (let i = 2; i < data.length; i++) { // skip 2 header rows
    const ohid = String(data[i][0]).trim(); // Column A
    if (!ohid) continue;

    const resolved = data[i][11]; // Column L checkbox (was M before column B was deleted)
    // All NYS orders are Baseline — no column B distinction needed
    const shouldCheck = !isCheckedResolved(resolved);

    records.push({ ohid, shouldCheck, typeValue: "Baseline" });
  }

  // xM NeXT Orders: OHID=B(2), Checkbox=F(6), NoteType=N(14), TypeD=C(3)
  syncCheckboxToMainTab(mainSheet, records, 6, 2, 14, 3);
  SpreadsheetApp.getUi().alert("NYS Audit sync complete. " + records.length + " records processed.");
}


// ============================================================
// BUTTON FUNCTION: Alerts Channel Audit
// OHIDs from col A. Checkbox col I.
// Col B determines Baseline/Subsequent for col D.
// Uncheck if col H checkbox is checked.
// ============================================================
function syncAlertsChannelAudit() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const auditSheet = ss.getSheetByName("Alerts Channel Audit");
  const mainSheet = ss.getSheetByName("xM NeXT Orders");

  if (!auditSheet || !mainSheet) {
    SpreadsheetApp.getUi().alert("Could not find required sheets. Please check sheet names.");
    return;
  }

  const data = auditSheet.getDataRange().getValues();
  const records = [];

  for (let i = 1; i < data.length; i++) {
    const ohid = String(data[i][0]).trim(); // Column A
    if (!ohid) continue;

    const resolved = data[i][7]; // Column H checkbox
    const type = String(data[i][1]).trim(); // Column B
    const typeValue = normalizeType(type);
    const shouldCheck = !isCheckedResolved(resolved);

    records.push({ ohid, shouldCheck, typeValue });
  }

  // xM NeXT Orders: OHID=B(2), Checkbox=H(8), NoteType=N(14), TypeD=C(3)
  syncCheckboxToMainTab(mainSheet, records, 8, 2, 14, 3);
  SpreadsheetApp.getUi().alert("Alerts Channel Audit sync complete. " + records.length + " records processed.");
}


// ============================================================
// BUTTON FUNCTION: Reports Delivered Sync
// OHIDs from col A. Notes to col N.
// Col B determines Baseline/Subsequent for col D.
// Blocks if duplicate OHIDs found in Reports Delivered.
// ============================================================
function syncReportsDelivered() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("Reports Delivered");
  const mainSheet = ss.getSheetByName("xM NeXT Orders");

  if (!sourceSheet || !mainSheet) {
    SpreadsheetApp.getUi().alert("Could not find required sheets. Please check sheet names.");
    return;
  }

  const sourceData = sourceSheet.getDataRange().getValues();
  const seen = {};
  const duplicates = new Set();
  const records = [];

  for (let i = 1; i < sourceData.length; i++) {
    const ohid = String(sourceData[i][0]).trim(); // Column A
    const type = String(sourceData[i][1]).trim(); // Column B
    const notes = sourceData[i][2];               // Column C

    if (!ohid) continue;

    if (seen[ohid]) {
      duplicates.add(ohid);
    } else {
      seen[ohid] = true;
    }

    records.push({ ohid, notes, typeValue: normalizeType(type) });
  }

  if (duplicates.size > 0) {
    const dupList = Array.from(duplicates).join(", ");
    SpreadsheetApp.getUi().alert(
      "Duplicate OHIDs found in 'Reports Delivered'. Please remove duplicates before running.\n\n" + dupList
    );
    return;
  }

  const NUM_COLS = 14;
  const lastRow = Math.max(mainSheet.getLastRow(), 1);
  const rawData = mainSheet.getRange(1, 1, lastRow, NUM_COLS).getValues();
  const data = rawData.map(row => {
    const padded = row.slice();
    while (padded.length < NUM_COLS) padded.push("");
    return padded;
  });

  const ohidMap = buildOhidMap(data, 2); // OHID in column B
  const newRows = [];

  records.forEach(({ ohid, notes, typeValue }) => {
    if (ohidMap[ohid] !== undefined) {
      const rowIndex = ohidMap[ohid];
      if (rowIndex < data.length) {
        if (isManualValue(data[rowIndex][14])) return; // respect Manual flag col O
        data[rowIndex][12] = notes; // Column M
        if (typeValue && !data[rowIndex][2]) {
          data[rowIndex][2] = typeValue; // Column C
        }
      }
    } else {
      const newRow = new Array(NUM_COLS).fill("");
      newRow[1] = ohid;       // Column B
      newRow[12] = notes;     // Column M
      newRow[13] = "Automatic"; // Column N
      if (typeValue) newRow[2] = typeValue; // Column C
      newRows.push(newRow);
      ohidMap[ohid] = data.length + newRows.length - 1;
    }
  });

  mainSheet.getRange(1, 1, data.length, NUM_COLS).setValues(data);

  if (newRows.length > 0) {
    const lastDataRow = mainSheet.getLastRow();
    let startRow = lastDataRow + 1;
    const ohidColData = mainSheet.getRange(2, 2, lastDataRow, 1).getValues();
    for (let i = 0; i < ohidColData.length; i++) {
      if (String(ohidColData[i][0]).trim() === "") {
        startRow = i + 2;
        break;
      }
    }
    mainSheet.getRange(startRow, 1, newRows.length, NUM_COLS).setValues(newRows);
  }

  SpreadsheetApp.getUi().alert(
    "Reports Delivered sync complete. " + records.length + " records processed."
  );
}


// ============================================================
// TRANSFER FUNCTION: Paste bulk data from xM CRC Paste sheet
// into xM CRC vs Next Audit
// ============================================================
function transferDataToCRC() {
  console.log('=== STARTING DATA TRANSFER ===');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName('xM CRC Paste sheet');
  const auditSheet = ss.getSheetByName('xM CRC vs Next Audit');

  if (!sourceSheet) {
    console.error('ERROR: "xM CRC Paste sheet" not found');
    SpreadsheetApp.getUi().alert('Error: "xM CRC Paste sheet" not found');
    return;
  }
  if (!auditSheet) {
    console.error('ERROR: "xM CRC vs Next Audit" not found');
    SpreadsheetApp.getUi().alert('Error: "xM CRC vs Next Audit" not found');
    return;
  }

  console.log('✓ Both sheets found');

  const sourceData = sourceSheet.getDataRange().getValues();
  const sourceRows = sourceData;

  console.log(`Source data: ${sourceRows.length} rows (no header in source)`);

  const auditData = auditSheet.getDataRange().getValues();

  console.log(`Audit data: ${auditData.length - 1} existing rows (excluding header)`);

  const existingColumnE = new Set();
  const existingColumnA = new Set();
  const idToAuditRow = {};
  const newRowIds = new Set();

  for (let i = 1; i < auditData.length; i++) {
    const colE = auditData[i][4]; // Column E (index 4)
    const colA = auditData[i][0]; // Column A (index 0)

    if (colE && String(colE).trim() !== '') {
      existingColumnE.add(String(colE).trim());
    }
    if (colA && String(colA).trim() !== '') {
      const colAStr = String(colA).trim();
      existingColumnA.add(colAStr);
      idToAuditRow[colAStr] = i + 1;
    }
  }

  console.log(`Found ${existingColumnE.size} unique values in Audit column E`);
  console.log(`Found ${existingColumnA.size} unique values in Audit column A`);

  const newRows = [];
  const skippedRows = [];
  let backfillCount = 0;

  console.log('\n=== PROCESSING SOURCE ROWS ===');

  for (let i = 0; i < sourceRows.length; i++) {
    const row = sourceRows[i];
    const sourceRowNum = i + 1;

    const colAValue = row[0] ? String(row[0]).trim() : ''; // Column A -> Audit A
    const colHValue = row[7] ? String(row[7]).trim() : ''; // Column H -> Audit E (duplicate check)

    console.log(`\n--- Row ${sourceRowNum} ---`);
    console.log(`  Column A: "${colAValue}", Column H: "${colHValue}"`);

    const colEExists = colHValue && existingColumnE.has(colHValue);
    const colAExists = colAValue && existingColumnA.has(colAValue);

    if (colEExists && colAExists) {
      if (newRowIds.has(colAValue)) {
        console.log(`  — SKIPPED Row ${sourceRowNum}: already added this run`);
        skippedRows.push(`Row ${sourceRowNum}: Duplicate — already added this run`);
        continue;
      }

      const auditRowNum = idToAuditRow[colAValue];

      if (!auditRowNum) {
        console.warn(`  ⚠ SKIPPED Row ${sourceRowNum}: could not find audit row for backfill`);
        skippedRows.push(`Row ${sourceRowNum}: Duplicate — could not find audit row, skipping backfill`);
        continue;
      }

      const auditRow = auditSheet.getRange(auditRowNum, 1, 1, 6).getValues()[0];
      const currentColB = auditRow[1];
      const sourceColB = row[1];

      if ((currentColB === '' || currentColB === null || currentColB === undefined) &&
          sourceColB !== '' && sourceColB !== null && sourceColB !== undefined) {
        auditSheet.getRange(auditRowNum, 2).setValue(sourceColB);
        console.log(`  ✓ BACKFILLED Row ${auditRowNum} Audit col B: "${sourceColB}"`);
        backfillCount++;
      } else {
        console.log(`  — SKIPPED Row ${sourceRowNum}: Duplicate, nothing to backfill`);
      }

      skippedRows.push(`Row ${sourceRowNum}: Duplicate in both Column A ("${colAValue}") and Column E ("${colHValue}") — ${currentColB === '' || currentColB === null ? 'backfilled B' : 'nothing to backfill'}`);
      continue;
    }

    console.log(`  ✓ ADDING row`);

    const newRow = [
      row[0],  // Column A (index 0) -> Audit A
      row[1],  // Column B (index 1) -> Audit B
      row[2],  // Column C (index 2) -> Audit C
      row[6],  // Column G (index 6) -> Audit D
      row[7],  // Column H (index 7) -> Audit E
      row[9],  // Column J (index 9) -> Audit F
    ];

    newRows.push(newRow);

    if (colAValue) {
      existingColumnA.add(colAValue);
      newRowIds.add(colAValue);
    }
    if (colHValue) existingColumnE.add(colHValue);
  }

  console.log(`\n=== PROCESSING COMPLETE ===`);
  console.log(`New rows to add: ${newRows.length}`);
  console.log(`Backfilled rows: ${backfillCount}`);
  console.log(`Skipped rows: ${skippedRows.length}`);

  let message = '';

  if (newRows.length > 0) {
    console.log('\n=== WRITING TO AUDIT SHEET ===');

    let lastRowWithData = 1;
    const columnAValues = auditSheet.getRange(1, 1, auditSheet.getMaxRows(), 1).getValues();

    for (let i = columnAValues.length - 1; i >= 0; i--) {
      if (columnAValues[i][0] !== '' && columnAValues[i][0] !== null && columnAValues[i][0] !== undefined) {
        lastRowWithData = i + 1;
        break;
      }
    }

    console.log(`Appending ${newRows.length} rows starting at row ${lastRowWithData + 1}`);

    const targetRange = auditSheet.getRange(lastRowWithData + 1, 1, newRows.length, 6);
    targetRange.setValues(newRows);

    console.log(`✓ Data written successfully`);

    message += `✓ Success! ${newRows.length} new row(s) transferred to xM CRC vs Next Audit.\n\n`;
  } else {
    console.warn('⚠ No new rows to transfer');
    message += `⚠ No new rows transferred.\n\n`;
  }

  if (backfillCount > 0) {
    message += `✓ Backfilled missing fields in ${backfillCount} existing row(s).\n\n`;
  }

  if (skippedRows.length > 0) {
    message += `Skipped ${skippedRows.length} row(s):\n`;
    message += skippedRows.slice(0, 10).join('\n');
    if (skippedRows.length > 10) {
      message += `\n... and ${skippedRows.length - 10} more`;
    }
  }

  console.log('\n=== TRANSFER COMPLETE ===');
  SpreadsheetApp.getUi().alert(message);
}

// ============================================================
// TRANSFER FUNCTION: Paste bulk data from Multiple Baseline
// Paste sheet into Multiple Baselines Audit
// ============================================================
function transferDataToMultipleBaselines() {
  console.log('=== STARTING DATA TRANSFER ===');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName('Multiple Baseline Paste sheet');
  const auditSheet = ss.getSheetByName('Multiple Baselines Audit');

  if (!sourceSheet) {
    console.error('ERROR: "Multiple Baseline Paste sheet" not found');
    SpreadsheetApp.getUi().alert('Error: "Multiple Baseline Paste sheet" not found');
    return;
  }
  if (!auditSheet) {
    console.error('ERROR: "Multiple Baselines Audit" not found');
    SpreadsheetApp.getUi().alert('Error: "Multiple Baselines Audit" not found');
    return;
  }

  console.log('✓ Both sheets found');

  const sourceData = sourceSheet.getDataRange().getValues();
  const sourceRows = sourceData;

  console.log(`Source data: ${sourceRows.length} rows (no header in source)`);

  const auditData = auditSheet.getDataRange().getValues();
  const existingIds = new Set();
  const idToAuditRow = {};
  const newRowIds = new Set();

  console.log(`Audit data: ${auditData.length - 1} existing rows (excluding header)`);

  for (let i = 1; i < auditData.length; i++) {
    const id = auditData[i][6];
    if (id && String(id).trim() !== '') {
      const idStr = String(id).trim();
      existingIds.add(idStr);
      idToAuditRow[idStr] = i + 1;
    }
  }

  console.log(`Found ${existingIds.size} unique IDs in Audit column G`);

  const newRows = [];
  const skippedRows = [];
  let backfillCount = 0;

  console.log('\n=== PROCESSING SOURCE ROWS ===');

  for (let i = 0; i < sourceRows.length; i++) {
    const row = sourceRows[i];
    const sourceRowNum = i + 1;

    const uniqueId = row[7] ? String(row[7]).trim() : '';

    if (!uniqueId || uniqueId === '') {
      console.warn(`  ⚠ SKIPPED Row ${sourceRowNum}: Empty ID in column H`);
      skippedRows.push(`Row ${sourceRowNum}: Empty ID in column H`);
      continue;
    }

    if (existingIds.has(uniqueId)) {
      if (newRowIds.has(uniqueId)) {
        console.log(`  — SKIPPED Row ${sourceRowNum}: Duplicate ID "${uniqueId}" already added this run`);
        skippedRows.push(`Row ${sourceRowNum}: Duplicate ID "${uniqueId}" — already added this run`);
        continue;
      }

      const auditRowNum = idToAuditRow[uniqueId];

      if (!auditRowNum) {
        console.warn(`  ⚠ SKIPPED Row ${sourceRowNum}: ID "${uniqueId}" found in existingIds but not in idToAuditRow — skipping backfill`);
        skippedRows.push(`Row ${sourceRowNum}: ID "${uniqueId}" — could not find audit row, skipping backfill`);
        continue;
      }

      const auditRow = auditSheet.getRange(auditRowNum, 1, 1, 12).getValues()[0];

      const backfills = [
        { auditCol: 3,  value: row[13] }, // N -> Audit C
        { auditCol: 4,  value: row[15] }, // P -> Audit D
        { auditCol: 9,  value: row[14] }, // O -> Audit I
        { auditCol: 10, value: row[16] }, // Q -> Audit J
      ];

      let rowBackfilled = false;
      for (const { auditCol, value } of backfills) {
        const currentVal = auditRow[auditCol - 1];
        if ((currentVal === '' || currentVal === null || currentVal === undefined) && value !== '' && value !== null && value !== undefined) {
          auditSheet.getRange(auditRowNum, auditCol).setValue(value);
          console.log(`  ✓ BACKFILLED Row ${auditRowNum} Audit col ${auditCol}: "${value}"`);
          rowBackfilled = true;
        }
      }

      if (rowBackfilled) backfillCount++;
      else console.log(`  — SKIPPED Row ${sourceRowNum}: ID "${uniqueId}" already exists, nothing to backfill`);

      skippedRows.push(`Row ${sourceRowNum}: Duplicate ID "${uniqueId}" — ${rowBackfilled ? 'backfilled missing fields' : 'nothing to backfill'}`);
      continue;
    }

    const newRow = [
      row[1],   // Column B (index 1)  -> Audit A
      row[2],   // Column C (index 2)  -> Audit B
      row[13],  // Column N (index 13) -> Audit C
      row[15],  // Column P (index 15) -> Audit D
      row[4],   // Column E (index 4)  -> Audit E
      row[5],   // Column F (index 5)  -> Audit F
      row[7],   // Column H (index 7)  -> Audit G (duplicate check)
      row[8],   // Column I (index 8)  -> Audit H
      row[14],  // Column O (index 14) -> Audit I
      row[16],  // Column Q (index 16) -> Audit J
      row[10],  // Column K (index 10) -> Audit K (date, used for sorting)
      row[11],  // Column L (index 11) -> Audit L
    ];

    console.log(`  ✓ ADDING Row ${sourceRowNum}: ID "${uniqueId}" | Date: ${row[10]}`);

    newRows.push(newRow);
    existingIds.add(uniqueId);
    newRowIds.add(uniqueId);
  }

  console.log(`\n=== PROCESSING COMPLETE ===`);
  console.log(`New rows to add: ${newRows.length}`);
  console.log(`Backfilled rows: ${backfillCount}`);
  console.log(`Skipped rows: ${skippedRows.length}`);

  let message = '';

  if (newRows.length > 0) {
    console.log('\n=== WRITING TO AUDIT SHEET ===');

    newRows.sort((a, b) => {
      const dateA = a[10];
      const dateB = b[10];
      if (!dateA && !dateB) return 0;
      if (!dateA) return 1;
      if (!dateB) return -1;
      return new Date(dateA).getTime() - new Date(dateB).getTime();
    });

    let lastRowWithData = 1;
    const columnAValues = auditSheet.getRange(1, 1, auditSheet.getMaxRows(), 1).getValues();

    for (let i = columnAValues.length - 1; i >= 0; i--) {
      if (columnAValues[i][0] !== '' && columnAValues[i][0] !== null && columnAValues[i][0] !== undefined) {
        lastRowWithData = i + 1;
        break;
      }
    }

    console.log(`Last row with data: ${lastRowWithData}`);
    console.log(`Appending ${newRows.length} rows starting at row ${lastRowWithData + 1}`);

    const targetRange = auditSheet.getRange(lastRowWithData + 1, 1, newRows.length, 12);
    targetRange.setValues(newRows);

    const sourceFormat = sourceSheet.getRange(1, 11, 1, 1).getNumberFormat();
    const dateColumn = auditSheet.getRange(lastRowWithData + 1, 11, newRows.length, 1);
    dateColumn.setNumberFormat(sourceFormat);

    console.log(`✓ Data written successfully`);

    message += `✓ Success! ${newRows.length} new row(s) transferred to Multiple Baselines Audit.\n`;
    message += `New data sorted by date (least recent to most recent).\n`;
    message += `Existing data was not modified.\n\n`;
  } else {
    console.warn('⚠ No new rows to transfer');
    message += `⚠ No new rows transferred.\n\n`;
  }

  if (backfillCount > 0) {
    message += `✓ Backfilled missing fields in ${backfillCount} existing row(s).\n\n`;
  }

  if (skippedRows.length > 0) {
    message += `Skipped ${skippedRows.length} row(s):\n`;
    message += skippedRows.slice(0, 10).join('\n');
    if (skippedRows.length > 10) {
      message += `\n... and ${skippedRows.length - 10} more`;
    }
  }

  console.log('\n=== TRANSFER COMPLETE ===');
  SpreadsheetApp.getUi().alert(message);
}


// ============================================================
// REMOVE DUPLICATES: Clean up duplicate rows in
// Multiple Baselines Audit based on column G ID
// ============================================================
function removeDuplicatesFromMultipleBaselines() {
  console.log('=== STARTING DUPLICATE REMOVAL ===');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const auditSheet = ss.getSheetByName('Multiple Baselines Audit');

  if (!auditSheet) {
    console.error('ERROR: "Multiple Baselines Audit" not found');
    SpreadsheetApp.getUi().alert('Error: "Multiple Baselines Audit" not found');
    return;
  }

  const auditData = auditSheet.getDataRange().getValues();
  const header = auditData[0];
  const dataRows = auditData.slice(1);

  const seenIds = new Set();
  const uniqueRows = [];
  let duplicateCount = 0;

  for (let i = 0; i < dataRows.length; i++) {
    const row = dataRows[i];
    const id = row[6] ? String(row[6]).trim() : '';

    if (id === '') {
      uniqueRows.push(row);
    } else if (!seenIds.has(id)) {
      seenIds.add(id);
      uniqueRows.push(row);
    } else {
      duplicateCount++;
      console.log(`Removing duplicate ID: "${id}" from row ${i + 2}`);
    }
  }

  if (duplicateCount > 0) {
    auditSheet.clear();
    auditSheet.getRange(1, 1, 1, header.length).setValues([header]);

    if (uniqueRows.length > 0) {
      auditSheet.getRange(2, 1, uniqueRows.length, uniqueRows[0].length).setValues(uniqueRows);
    }

    console.log(`✓ Removed ${duplicateCount} duplicate row(s)`);
    SpreadsheetApp.getUi().alert(`✓ Removed ${duplicateCount} duplicate row(s) from Multiple Baselines Audit.\n${uniqueRows.length} unique rows remaining.`);
  } else {
    console.log('No duplicates found');
    SpreadsheetApp.getUi().alert('No duplicates found in Multiple Baselines Audit.');
  }

  console.log('=== DUPLICATE REMOVAL COMPLETE ===');
}

// ============================================================
// BUTTON FUNCTION: Initialize manually-added OHIDs
// One-time, fill-only sweep of xM NeXT Orders.
// For every row with an OHID in column B, fills A="Active" and
// N="Automatic" ONLY where those cells are currently blank.
// Never overwrites an existing A or N value, so prior rows keep
// their status / flag exactly as-is.
// ============================================================
function initializeManualOhids() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("xM NeXT Orders");
  if (!sheet) {
    SpreadsheetApp.getUi().alert("Could not find 'xM NeXT Orders'. Please check the sheet name.");
    return;
  }

  const NUM_COLS = 14;
  const lastRow = Math.max(sheet.getLastRow(), 1);
  const data = sheet.getRange(1, 1, lastRow, NUM_COLS).getValues();

  let updated = 0;
  for (let i = 1; i < data.length; i++) { // skip header
    const ohid = String(data[i][1]).trim(); // Column B
    if (!ohid) continue;

    let changed = false;
    if (String(data[i][0]).trim() === "") { // Column A only if blank
      data[i][0] = "Active";
      changed = true;
    }
    if (String(data[i][13]).trim() === "") { // Column N only if blank
      data[i][13] = "Automatic";
      changed = true;
    }
    if (changed) updated++;
  }

  sheet.getRange(1, 1, data.length, NUM_COLS).setValues(data);
  SpreadsheetApp.getUi().alert("Initialized " + updated + " row(s) with Active / Automatic.");
}


// ============================================================
// SETUP: Run this function ONCE to install the onEdit trigger
// After running, delete or ignore this function
// ============================================================
function installOnEditTrigger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Remove any existing onEditTrigger triggers to avoid duplicates
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === "onEditTrigger") {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  // Install fresh trigger
  ScriptApp.newTrigger("onEditTrigger")
    .forSpreadsheet(ss)
    .onEdit()
    .create();
  SpreadsheetApp.getUi().alert("onEdit trigger installed successfully.");
}
