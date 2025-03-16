function sendDataToSheets(selectedItems) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Lesson Plan Data");

  // If the sheet doesn't exist, create it
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Lesson Plan Data");
  }

  var lastRow = sheet.getLastRow();
  var nextRow = lastRow + 1;

  // Insert the selected checkboxes in the next available row
  sheet.getRange(nextRow, 1, 1, selectedItems.length).setValues([selectedItems]);
}


function showPopup() {
  var html = HtmlService.createHtmlOutputFromFile('PopupMenu')
    .setWidth(300)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, "Select Items");
}
