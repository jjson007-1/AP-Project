/*function exportToGoogleDocs() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Plan"); // Change to your sheet name
  var folderId = "1trrmSlwPW4swXk3VyVM2CKL0VMOgSNez"; // Replace with your Google Drive folder ID
  var docIdCell = sheet.getRange("A1"); // Store the Google Doc ID in cell A1
  var docId = docIdCell.getValue();
  var doc;

  if (!docId) {
    // If no Doc ID is stored, create a new Google Doc
    var folder = DriveApp.getFolderById(folderId);
    doc = DocumentApp.create("Exported Data from " + sheet.getName());
    var newDocId = doc.getId();
    folder.addFile(DriveApp.getFileById(newDocId)); // Move to the target folder

    // Store the new Doc ID in the sheet
    docIdCell.setValue(newDocId);
  } else {
    // Open existing document
    doc = DocumentApp.openById(docId);
  }

  var body = doc.getBody();
  var data = sheet.getDataRange().getValues();

  // Clear existing content (optional)
  body.clear();

  // Insert new data
  for (var i = 0; i < data.length; i++) {
    body.appendParagraph(data[i].join(" | ")); // Formats data with ' | ' separator
  }

  doc.saveAndClose();
}*/

/*function exportToGoogleDocs() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Plan"); // Change to your sheet name
  var folderId = "1trrmSlwPW4swXk3VyVM2CKL0VMOgSNez"; // Replace with your Google Drive folder ID
  var docIdCell = sheet.getRange("A1"); // Store the Google Doc ID in cell A1
  var docId = docIdCell.getValue();
  var doc;

  if (!docId) {
    // If no Doc ID is stored, create a new Google Doc
    var folder = DriveApp.getFolderById(folderId);
    doc = DocumentApp.create("Exported Data from " + sheet.getName());
    var newDocId = doc.getId();
    folder.addFile(DriveApp.getFileById(newDocId)); // Move to the target folder

    // Store the new Doc ID in the sheet
    docIdCell.setValue(newDocId);
  } else {
    // Open existing document
    doc = DocumentApp.openById(docId);
  }

  var body = doc.getBody();
  var data = sheet.getRange("A2:B6").getValues(); // Change this to your desired range

  // Clear existing content (optional)
  //body.clear();

  // Insert new data
  for (var i = 0; i < data.length; i++) {
    body.appendParagraph(data[i].join(" ")); // Joins row data with spaces, each row on a new line
  }

  doc.saveAndClose();
}*/

/*function exportToNewGoogleDoc() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet20"); // Change to your sheet name
  var folderId = "1trrmSlwPW4swXk3VyVM2CKL0VMOgSNez"; // Replace with your Google Drive folder ID
  var folder = DriveApp.getFolderById(folderId); // Get the target folder
  
  // Create a new Google Document with a timestamp in the name
  var docName = sheet.getRange("B1").getValue();
  var doc = DocumentApp.create(docName + " - " + new Date().toLocaleString());
  var docId = doc.getId();
  folder.addFile(DriveApp.getFileById(docId)); // Move to the target folder

  var body = doc.getBody();
  var data = sheet.getRange("D7").getValues(); // Change this to your desired range

  // Insert new data into the new document
  for (var i = 0; i < data.length; i++) {
    body.appendParagraph(data[i].join(" ")); // Joins row data with spaces, each row on a new line
  }

  doc.saveAndClose();

  // (Optional) Store the new Document ID in Google Sheets for reference
  sheet.getRange("A1").setValue(docId);

  // (Optional) Show a success message
  SpreadsheetApp.getUi().alert("Data exported to new Google Doc:\n" + doc.getUrl());
}*/

function exportToNewGoogleDoc() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet20");
  var sheet2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Plan");
  var folderId = "1trrmSlwPW4swXk3VyVM2CKL0VMOgSNez";
  var folder = DriveApp.getFolderById(folderId);

  // Create new document with timestamp
  var docName = sheet.getRange("B1").getValue();
  var doc = DocumentApp.create(docName + " - " + new Date().toLocaleString());
  var docId = doc.getId();
  folder.addFile(DriveApp.getFileById(docId));

  var body = doc.getBody();

  // Get data from D7
  var dataMain = sheet.getRange("D7").getValues();

  // Insert main data
  body.appendParagraph("Main Plan:");
  for (var i = 0; i < dataMain.length; i++) {
    body.appendParagraph(dataMain[i].join(" "));
  }

  // Add a page break
  body.appendPageBreak();

  // Add title for Detailed Plan
  body.appendParagraph("Detailed Plan").setHeading(DocumentApp.ParagraphHeading.HEADING1);

  // Get additional data (adjust the range as needed)
  var detailedData = sheet2.getRange("B2:B6").getValues();

  // Insert detailed data
  for (var j = 0; j < detailedData.length; j++) {
    body.appendParagraph(detailedData[j][0]); // Single column data
  }

  doc.saveAndClose();

  // Store the new Document ID
  sheet.getRange("A1").setValue(docId);

  // Show a success message
  SpreadsheetApp.getUi().alert("Data exported to new Google Doc");
}




