
function copyAndSortAllDatesChronologically() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("Sheet to paste");
  const targetSheet = ss.getSheetByName("Sheet1");

  if (!sourceSheet || !targetSheet) {
    throw new Error('Either "Sheet to paste" or "Sheet1" does not exist. Please check sheet names.');
  }

  const headerRows = 1;
  const data = sourceSheet.getDataRange().getValues();

  // Column indexes (0-based)
  const colE = 4, colF = 5, colG = 6, colI = 8, colM = 12, colN = 13, colO = 14, colP = 15, colQ = 16;

  // --- Get existing data from Sheet1 to check for duplicates ---
  const existingData = targetSheet.getDataRange().getValues();
  const existingKeys = new Set();

  if (existingData.length > 0) {
    existingData.forEach(row => {
      const key = `${row[0]}|${row[1]}|${row[2]}|${row[3]}`;
      existingKeys.add(key);
    });
  }

  // --- Extract valid rows from source ---
  const bodyData = data.slice(headerRows).filter(r => r[colN]);
  if (bodyData.length === 0) {
    Logger.log("No rows with a date found in column N.");
    return;
  }

  // --- Helper: convert to Date safely ---
  const toDate = (v) => {
    if (v instanceof Date) return new Date(v.getFullYear(), v.getMonth(), v.getDate());
    const parsed = new Date(v);
    if (!isNaN(parsed)) return new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate());
    const parts = v.toString().split(/[-/]/);
    if (parts.length >= 3) {
      const [p1, p2, p3] = parts.map(Number);
      const y = p3 < 100 ? 2000 + p3 : p3;
      const m = p1 - 1;
      const d = p2;
      return new Date(y, m, d);
    }
    return new Date("2100-01-01");
  };

  // --- Build a lookup map from source: key → { colO, colP } ---
  const sourceOPMap = {};
  bodyData.forEach(row => {
    const dateObj = toDate(row[colN]);
    const uniqueKey = `${dateObj}|${row[colE]}|${row[colF]}|${row[colG]}`;
    if (!sourceOPMap[uniqueKey]) {
      sourceOPMap[uniqueKey] = { o: row[colO], p: row[colP] };
    }
  });

  // --- Second pass: fill blank G/H in existing Sheet1 rows ---
  let filledCount = 0;
  if (existingData.length > 0) {
    existingData.forEach((row, i) => {
      const gIsEmpty = row[6] === "" || row[6] === null || row[6] === undefined;
      const hIsEmpty = row[7] === "" || row[7] === null || row[7] === undefined;

      if (gIsEmpty || hIsEmpty) {
        const key = `${row[0]}|${row[1]}|${row[2]}|${row[3]}`;
        const match = sourceOPMap[key];
        if (match) {
          const sheetRow = i + 1; // 1-based
          if (gIsEmpty && match.o !== undefined) {
            targetSheet.getRange(sheetRow, 7).setValue(match.o);
          }
          if (hIsEmpty && match.p !== undefined) {
            targetSheet.getRange(sheetRow, 8).setValue(match.p);
          }
          filledCount++;
        }
      }
    });
  }

  // --- Group rows by date, filtering duplicates ---
  const groupedByDate = {};
  let skippedDuplicates = 0;

  bodyData.forEach(row => {
    const dateObj = toDate(row[colN]);
    const uniqueKey = `${dateObj}|${row[colE]}|${row[colF]}|${row[colG]}`;

    if (existingKeys.has(uniqueKey)) {
      skippedDuplicates++;
      return;
    }

    const key = dateObj.getTime();
    if (!groupedByDate[key]) groupedByDate[key] = [];
    groupedByDate[key].push({ row, dateObj });
  });

  // --- Sort date groups chronologically ---
  const sortedDateKeys = Object.keys(groupedByDate).map(Number).sort((a, b) => a - b);

  // --- Build final output ---
  let output = [];
  sortedDateKeys.forEach(dateKey => {
    const entries = groupedByDate[dateKey];

    entries.sort((a, b) => {
      const valA = a.row[colE] ? a.row[colE].toString().toLowerCase() : "";
      const valB = b.row[colE] ? b.row[colE].toString().toLowerCase() : "";
      return valA.localeCompare(valB);
    });

    const mapped = entries.map(e => [
      e.dateObj,       // → Sheet1 col A
      e.row[colE],     // → Sheet1 col B
      e.row[colF],     // → Sheet1 col C
      e.row[colG],     // → Sheet1 col D
      e.row[colI],     // → Sheet1 col E
      e.row[colM],     // → Sheet1 col F
      e.row[colO],     // → Sheet1 col G
      e.row[colP],     // → Sheet1 col H
      e.row[colQ],     // → Sheet1 col I
    ]);

    output = output.concat(mapped);
  });

  // --- Append only new rows to Sheet1 ---
  if (output.length > 0) {
    const lastRow = targetSheet.getLastRow();
    const startRow = lastRow > 0 ? lastRow + 1 : 1;

    targetSheet
      .getRange(startRow, 1, output.length, output[0].length)
      .setValues(output);

    targetSheet.getRange(startRow, 1, output.length, 1).setNumberFormat("MM/dd/yy");
  }

  Logger.log(`Appended ${output.length} new rows. Skipped ${skippedDuplicates} duplicates. Filled G/H for ${filledCount} existing rows.`);
}