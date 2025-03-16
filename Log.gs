function logAndEmail() {
  var email = "umbrelytics@gmail.com"; // Replace with your email
  var subject = "Google Sheets Script Log Data";

  try {
    // Example: Simulate an error (change this to your real function)
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("NonExistentSheet");
    var data = sheet.getRange("A1").getValue(); 

  } catch (e) {
    // Capture error and send it via email
    var message = "An error occurred in your script:\n\n" + e.toString();
    MailApp.sendEmail(email, subject, message);
  }
}
