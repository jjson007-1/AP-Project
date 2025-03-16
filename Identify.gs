var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
function matchAndPrintMultipleRangesIgnoreSpaces() {
  // Define the target cell to match (ignore spaces and case)
  var targetCell = sheet.getRange("E7").getValue().replace(/\s+/g, "").toLowerCase();

  // Define 4 search ranges and their corresponding result ranges
  var ranges = [
    { search: sheet.getRange("Resources!AA2:AA11").getValues(), result1: sheet.getRange("Resources!AB2:AB11").getValues(), result2: sheet.getRange("Resources!AC2:AC11").getValues()},
    { search: sheet.getRange("Resources!AD2:AD7").getValues(), result1: sheet.getRange("Resources!AE2:AE7").getValues(), result2: sheet.getRange("Resources!AF2:AF7").getValues() },
    { search: sheet.getRange("Resources!AG2:AG9").getValues(), result1: sheet.getRange("Resources!AH2:AH9").getValues(), result2: sheet.getRange("Resources!AI2:AI9").getValues() },
    { search: sheet.getRange("Resources!AJ2:AJ9").getValues(), result1: sheet.getRange("Resources!AK2:AK9").getValues(), result2: sheet.getRange("Resources!AL2:AL9").getValues() }  
  ];

  // Initialize result variable
  var result1 = "";
  var result2 = "";

  // Loop through each range
  for (var i = 0; i < ranges.length; i++) {
    var searchRange = ranges[i].search;
    var resultRange1 = ranges[i].result1;
    var resultRange2 = ranges[i].result2;

    // Check each cell in the current search range
    for (var j = 0; j < searchRange.length; j++) {
      var currentValue = searchRange[j][0].replace(/\s+/g, "").toLowerCase(); // Remove spaces and convert to lowercase

      if (currentValue === targetCell) {
        result1 = resultRange1[j][0]; // Get the corresponding result value
        result2 = resultRange2[j][0]; // Get the corresponding result value
        break; // Stop searching once a match is found
      }
    }

    if (result1 && result2) break; // Stop searching through other ranges if a match is found
  }

  // Print the result in a specific cell
  sheet.getRange("F7").setValue(result1 ? result1 : ""); // If no match, display ""
  sheet.getRange("G7").setValue(result2 ? result2 : ""); // If no match, display ""

}


function matchAndPrintMultipleRangesIgnoreSpaces2() {
  //var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Define the target cell to match (ignore spaces and case)
  var targetCell = sheet.getRange("O7").getValue().replace(/\s+/g, "").toLowerCase();

  // Define 4 search ranges and their corresponding result ranges
  var ranges = [
    { search: sheet.getRange("Resources!AA2:AA11").getValues(), result1: sheet.getRange("Resources!AB2:AB11").getValues(), result2: sheet.getRange("Resources!AC2:AC11").getValues()},
    { search: sheet.getRange("Resources!AD2:AD7").getValues(), result1: sheet.getRange("Resources!AE2:AE7").getValues(), result2: sheet.getRange("Resources!AF2:AF7").getValues() },
    { search: sheet.getRange("Resources!AG2:AG9").getValues(), result1: sheet.getRange("Resources!AH2:AH9").getValues(), result2: sheet.getRange("Resources!AI2:AI9").getValues() },
    { search: sheet.getRange("Resources!AJ2:AJ9").getValues(), result1: sheet.getRange("Resources!AK2:AK9").getValues(), result2: sheet.getRange("Resources!AL2:AL9").getValues() }  
  ];

  // Initialize result variable
  var result1 = "";
  var result2 = "";

  // Loop through each range
  for (var i = 0; i < ranges.length; i++) {
    var searchRange = ranges[i].search;
    var resultRange1 = ranges[i].result1;
    var resultRange2 = ranges[i].result2;

    // Check each cell in the current search range
    for (var j = 0; j < searchRange.length; j++) {
      var currentValue = searchRange[j][0].replace(/\s+/g, "").toLowerCase(); // Remove spaces and convert to lowercase

      if (currentValue === targetCell) {
        result1 = resultRange1[j][0]; // Get the corresponding result value
        result2 = resultRange2[j][0]; // Get the corresponding result value
        break; // Stop searching once a match is found
      }
    }

    if (result1 && result2) break; // Stop searching through other ranges if a match is found
  }

  // Print the result in a specific cell
  sheet.getRange("P7").setValue(result1 ? result1 : ""); // If no match, display ""
  sheet.getRange("Q7").setValue(result2 ? result2 : ""); // If no match, display ""

}


function matchAndPrintMultipleRangesIgnoreSpaces3() {
  //var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Define the target cell to match (ignore spaces and case)
  var targetCell = sheet.getRange("Y7").getValue().replace(/\s+/g, "").toLowerCase();

  // Define 4 search ranges and their corresponding result ranges
  var ranges = [
    { search: sheet.getRange("Resources!AA2:AA11").getValues(), result1: sheet.getRange("Resources!AB2:AB11").getValues(), result2: sheet.getRange("Resources!AC2:AC11").getValues()},
    { search: sheet.getRange("Resources!AD2:AD7").getValues(), result1: sheet.getRange("Resources!AE2:AE7").getValues(), result2: sheet.getRange("Resources!AF2:AF7").getValues() },
    { search: sheet.getRange("Resources!AG2:AG9").getValues(), result1: sheet.getRange("Resources!AH2:AH9").getValues(), result2: sheet.getRange("Resources!AI2:AI9").getValues() },
    { search: sheet.getRange("Resources!AJ2:AJ9").getValues(), result1: sheet.getRange("Resources!AK2:AK9").getValues(), result2: sheet.getRange("Resources!AL2:AL9").getValues() }  
  ];

  // Initialize result variable
  var result1 = "";
  var result2 = "";

  // Loop through each range
  for (var i = 0; i < ranges.length; i++) {
    var searchRange = ranges[i].search;
    var resultRange1 = ranges[i].result1;
    var resultRange2 = ranges[i].result2;

    // Check each cell in the current search range
    for (var j = 0; j < searchRange.length; j++) {
      var currentValue = searchRange[j][0].replace(/\s+/g, "").toLowerCase(); // Remove spaces and convert to lowercase

      if (currentValue === targetCell) {
        result1 = resultRange1[j][0]; // Get the corresponding result value
        result2 = resultRange2[j][0]; // Get the corresponding result value
        break; // Stop searching once a match is found
      }
    }

    if (result1 && result2) break; // Stop searching through other ranges if a match is found
  }

  // Print the result in a specific cell
  sheet.getRange("Z7").setValue(result1 ? result1 : ""); // If no match, display ""
  sheet.getRange("AA7").setValue(result2 ? result2 : ""); // If no match, display ""

}


function matchAndPrintMultipleRangesIgnoreSpaces4() {
  //var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Define the target cell to match (ignore spaces and case)
  var targetCell = sheet.getRange("AI7").getValue().replace(/\s+/g, "").toLowerCase();

  // Define 4 search ranges and their corresponding result ranges
  var ranges = [
    { search: sheet.getRange("Resources!AA2:AA11").getValues(), result1: sheet.getRange("Resources!AB2:AB11").getValues(), result2: sheet.getRange("Resources!AC2:AC11").getValues()},
    { search: sheet.getRange("Resources!AD2:AD7").getValues(), result1: sheet.getRange("Resources!AE2:AE7").getValues(), result2: sheet.getRange("Resources!AF2:AF7").getValues() },
    { search: sheet.getRange("Resources!AG2:AG9").getValues(), result1: sheet.getRange("Resources!AH2:AH9").getValues(), result2: sheet.getRange("Resources!AI2:AI9").getValues() },
    { search: sheet.getRange("Resources!AJ2:AJ9").getValues(), result1: sheet.getRange("Resources!AK2:AK9").getValues(), result2: sheet.getRange("Resources!AL2:AL9").getValues() }  
  ];

  // Initialize result variable
  var result1 = "";
  var result2 = "";

  // Loop through each range
  for (var i = 0; i < ranges.length; i++) {
    var searchRange = ranges[i].search;
    var resultRange1 = ranges[i].result1;
    var resultRange2 = ranges[i].result2;

    // Check each cell in the current search range
    for (var j = 0; j < searchRange.length; j++) {
      var currentValue = searchRange[j][0].replace(/\s+/g, "").toLowerCase(); // Remove spaces and convert to lowercase

      if (currentValue === targetCell) {
        result1 = resultRange1[j][0]; // Get the corresponding result value
        result2 = resultRange2[j][0]; // Get the corresponding result value
        break; // Stop searching once a match is found
      }
    }

    if (result1 && result2) break; // Stop searching through other ranges if a match is found
  }

  // Print the result in a specific cell
  sheet.getRange("AJ7").setValue(result1 ? result1 : ""); // If no match, display ""
  sheet.getRange("AK7").setValue(result2 ? result2 : ""); // If no match, display ""

}


function matchAndPrintMultipleRangesIgnoreSpaces5() {
  //var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Define the target cell to match (ignore spaces and case)
  var targetCell = sheet.getRange("AS7").getValue().replace(/\s+/g, "").toLowerCase();

  // Define 4 search ranges and their corresponding result ranges
  var ranges = [
    { search: sheet.getRange("Resources!AA2:AA11").getValues(), result1: sheet.getRange("Resources!AB2:AB11").getValues(), result2: sheet.getRange("Resources!AC2:AC11").getValues()},
    { search: sheet.getRange("Resources!AD2:AD7").getValues(), result1: sheet.getRange("Resources!AE2:AE7").getValues(), result2: sheet.getRange("Resources!AF2:AF7").getValues() },
    { search: sheet.getRange("Resources!AG2:AG9").getValues(), result1: sheet.getRange("Resources!AH2:AH9").getValues(), result2: sheet.getRange("Resources!AI2:AI9").getValues() },
    { search: sheet.getRange("Resources!AJ2:AJ9").getValues(), result1: sheet.getRange("Resources!AK2:AK9").getValues(), result2: sheet.getRange("Resources!AL2:AL9").getValues() }  
  ];

  // Initialize result variable
  var result1 = "";
  var result2 = "";

  // Loop through each range
  for (var i = 0; i < ranges.length; i++) {
    var searchRange = ranges[i].search;
    var resultRange1 = ranges[i].result1;
    var resultRange2 = ranges[i].result2;

    // Check each cell in the current search range
    for (var j = 0; j < searchRange.length; j++) {
      var currentValue = searchRange[j][0].replace(/\s+/g, "").toLowerCase(); // Remove spaces and convert to lowercase

      if (currentValue === targetCell) {
        result1 = resultRange1[j][0]; // Get the corresponding result value
        result2 = resultRange2[j][0]; // Get the corresponding result value
        break; // Stop searching once a match is found
      }
    }

    if (result1 && result2) break; // Stop searching through other ranges if a match is found
  }

  // Print the result in a specific cell
  sheet.getRange("AT7").setValue(result1 ? result1 : ""); // If no match, display ""
  sheet.getRange("AU7").setValue(result2 ? result2 : ""); // If no match, display ""

}


function matchAndPrintMultipleRangesIgnoreSpaces6() {
  //var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Define the target cell to match (ignore spaces and case)
  var targetCell = sheet.getRange("BC7").getValue().replace(/\s+/g, "").toLowerCase();

  // Define 4 search ranges and their corresponding result ranges
  var ranges = [
    { search: sheet.getRange("Resources!AA2:AA11").getValues(), result1: sheet.getRange("Resources!AB2:AB11").getValues(), result2: sheet.getRange("Resources!AC2:AC11").getValues()},
    { search: sheet.getRange("Resources!AD2:AD7").getValues(), result1: sheet.getRange("Resources!AE2:AE7").getValues(), result2: sheet.getRange("Resources!AF2:AF7").getValues() },
    { search: sheet.getRange("Resources!AG2:AG9").getValues(), result1: sheet.getRange("Resources!AH2:AH9").getValues(), result2: sheet.getRange("Resources!AI2:AI9").getValues() },
    { search: sheet.getRange("Resources!AJ2:AJ9").getValues(), result1: sheet.getRange("Resources!AK2:AK9").getValues(), result2: sheet.getRange("Resources!AL2:AL9").getValues() }  
  ];

  // Initialize result variable
  var result1 = "";
  var result2 = "";

  // Loop through each range
  for (var i = 0; i < ranges.length; i++) {
    var searchRange = ranges[i].search;
    var resultRange1 = ranges[i].result1;
    var resultRange2 = ranges[i].result2;

    // Check each cell in the current search range
    for (var j = 0; j < searchRange.length; j++) {
      var currentValue = searchRange[j][0].replace(/\s+/g, "").toLowerCase(); // Remove spaces and convert to lowercase

      if (currentValue === targetCell) {
        result1 = resultRange1[j][0]; // Get the corresponding result value
        result2 = resultRange2[j][0]; // Get the corresponding result value
        break; // Stop searching once a match is found
      }
    }

    if (result1 && result2) break; // Stop searching through other ranges if a match is found
  }

  // Print the result in a specific cell
  sheet.getRange("BD7").setValue(result1 ? result1 : ""); // If no match, display ""
  sheet.getRange("BE7").setValue(result2 ? result2 : ""); // If no match, display ""

}



/*function addSkillMatchMenu(ui) {
  ui.createMenu('Skill Match')
    .addItem('Match Skills', 'matchAndPrintMultipleRangesIgnoreSpaces')  
    .addToUi();
}*/


