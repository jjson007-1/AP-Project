function openReportForm() {
 var ss = SpreadsheetApp.getActiveSpreadsheet();
  var range = ss.getRange('B1:E1'); 
  var name = range.getValues()[0].join(' '); 

  var html = HtmlService.createHtmlOutputFromFile('ReportForm')
  .setWidth(400)
  .setHeight(400)
  .append('<script>console.log("Prefill Name: ' + name + '"); var prefillName = "' + name + '";</script>');
  SpreadsheetApp.getUi().showModalDialog(html, "Submit Report");
}

function sendReport(form) {
  const email = "umbrelytics@gmail.com";
  const subject = "New Report Submission"
  const message = `Category: ${form.category}\nReported By: ${form.name}\nEmail: ${form.email}\nDescription: ${form.message}`;

  MailApp.sendEmail(email, subject, message);

  return "Report Submitted! Thank you for your feedback.";
}

