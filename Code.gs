/**
 * DISCLAIMER:
 * This script uses SpreadsheetApp.getUi().alert() to show messages.
 * ⚠️ Execution will pause until you confirm each prompt inside Google Sheets.
 * If you run this script from the Apps Script editor, make sure the sheet is open
 * and you click OK on the dialogs — otherwise the script will appear to "hang."
 */

function writeSheetFile() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 👇 Modify this to the sheet you want to export from
  const sourceSheetName = "Sheet1"; 
  const sourceSheet = ss.getSheetByName(sourceSheetName);
  if (!sourceSheet) {
    SpreadsheetApp.getUi().alert(`Source sheet "${sourceSheetName}" not found!`);
    return;
  }

  // Get all data from source sheet
  const data = sourceSheet.getDataRange().getValues();
  if (data.length === 0) {
    SpreadsheetApp.getUi().alert(`No data found in "${sourceSheetName}"!`);
    return;
  }

  // 👇 Modify this to customize the base name of the exported files
  const baseFileName = "Exported_Data"; 
  const timestamp = new Date().toISOString().replace(/[:.]/g, "-");
  const fileName = `${baseFileName}_${timestamp}`;

  // Create a new Google Sheet
  const newSpreadsheet = SpreadsheetApp.create(fileName);
  const newSheet = newSpreadsheet.getActiveSheet();

  // Paste data
  newSheet.getRange(1, 1, data.length, data[0].length).setValues(data);

  // Save as Excel file
  const blob = newSpreadsheet.getBlob().setName(fileName + ".xlsx");
  DriveApp.createFile(blob);

  SpreadsheetApp.getUi().alert(`Data exported to new sheet: "${fileName}"`);
}
