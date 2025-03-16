/*function savePlan() {
  // Get the active spreadsheet
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // Get the "Outline" sheet
  var sheet = spreadsheet.getSheetByName("Plan");
  if (!sheet) {
    throw new Error("Sheet 'Outline' not found.");
  }

  // Get the "Lesson Plan History" sheet
  var resultSheet = spreadsheet.getSheetByName("Lesson Plan History");
  if (!resultSheet) {
    throw new Error("Sheet 'Lesson Plan History' not found.");
  }

  // Get the target cell from "Outline"
  var targetCell = sheet.getRange("B2:B6"); 
  var result = targetCell.getValues(); // Get value from A3 in "Outline"

  // Find the next available row in "Lesson Plan History", starting from A2
  var lastRow = resultSheet.getLastRow(); // Get last non-empty row
  var nextRow = lastRow >= 2 ? lastRow + 2 : 2; // Start from A2, move down in steps of 2

  // Insert value in the next available row in Column A
  resultSheet.getRange(nextRow, 1).setValue(result);
}*/

function savePlan() {
  // Get the active spreadsheet
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // Get the "Plan" sheet
  var sheet = spreadsheet.getSheetByName("Plan");
  if (!sheet) {
    throw new Error("Sheet 'Plan' not found.");
  }

  // Get the "Lesson Plan History" sheet
  var resultSheet = spreadsheet.getSheetByName("Lesson Plan History");
  if (!resultSheet) {
    throw new Error("Sheet 'Lesson Plan History' not found.");
  }

  // Get the target range from "Plan" (B2:B6)
  var targetRange = sheet.getRange("B2:B6"); 
  var data = targetRange.getValues(); // Get values as a 2D array

  // Find the next available row in "Lesson Plan History"
  var lastRow = resultSheet.getLastRow();
  var nextRow = lastRow >= 1 ? lastRow + 3 : 1; // Start from row 1 if empty

  // Paste the entire range starting from the next available row
  resultSheet.getRange(nextRow, 1, data.length, data[0].length).setValues(data);

  // Add two blank rows after the pasted data
  var blankRows = 2;
  resultSheet.insertRowsAfter(nextRow + data.length - 1, blankRows);
}



