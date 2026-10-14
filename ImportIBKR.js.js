/**
 * IBKR Phase 1 – full-width CSV import.
 * Writes every column from the Flex CSV into the Import sheet.
 * Skips blank rows and any row that looks like a repeated header.
 */
function importIBKRFromFolder() {
  const FOLDER_ID = '17BdL5ipSYavF4uRlJdcy9l4nW9azWeln'; // ← your folder ID
  const sheet = SpreadsheetApp.getActive().getSheetByName('Import');
  if (!sheet) throw new Error('Import sheet not found');

  // Clear old data but keep the header row
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  }

  const folder = DriveApp.getFolderById(FOLDER_ID);
  const files = folder.getFilesByType(MimeType.CSV);
  const allData = [];

  while (files.hasNext()) {
    const file = files.next();
    const csv = file.getBlob().getDataAsString();
    const rows = Utilities.parseCsv(csv);

    for (let i = 1; i < rows.length; i++) {          // skip file header
      const row = rows[i];
      if (!row || row.length === 0) continue;

      // Skip repeated header blocks and pure-blank rows
      const firstCell = String(row[0] || '').trim();
      if (firstCell === 'ClientAccountID' || firstCell === '') continue;

      // Pad or trim so every row matches the sheet header width
      const headerWidth = 45; // number of columns in the header above
      while (row.length < headerWidth) row.push('');
      if (row.length > headerWidth) row.length = headerWidth;

      allData.push(row);
    }
  }

  if (allData.length === 0) {
    SpreadsheetApp.getUi().alert('No data rows found');
    return;
  }

  sheet.getRange(2, 1, allData.length, allData[0].length).setValues(allData);
  SpreadsheetApp.getUi().alert('Imported ' + allData.length + ' full-width IBKR rows');
}