var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet1 = spreadsheet.getSheetByName("Lesson Plan AI Prompt"); // Replace with the actual name of your first sheet
  var sheet2 = spreadsheet.getSheetByName("Skills Times"); // Replace with the actual name of your second sheet
  var sheet3 = spreadsheet.getSheetByName("Lesson Planning");

function matchAndPrintFromTwoSheets() {
  // Get the spreadsheet and sheets
  /*var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet1 = spreadsheet.getSheetByName("Lesson Plan AI Prompt"); // Replace with the actual name of your first sheet
  var sheet2 = spreadsheet.getSheetByName("Skills Times"); // Replace with the actual name of your second sheet
  var sheet3 = spreadsheet.getSheetByName("Lesson Planning");*/

  if (!sheet1 || !sheet2 || !sheet3) {
    throw new Error("One or both sheets are not found. Ensure sheet names are correct.");
  }

  // Get values from the first sheet
  var valueA1 = sheet1.getRange("D2").getValue().toString().trim().toLowerCase();
  var valueB1 = sheet1.getRange("C2").getValue().toString().trim().toLowerCase();
  var level = sheet3.getRange("B3").getValue().toString().trim().toLowerCase();

  // Define 4 search ranges and their corresponding left and right ranges
  var ranges = [
    {
      search: sheet2.getRange("B2:B12").getValues(),
      left: sheet2.getRange("A2:A12").getValues(),
      right: sheet2.getRange("C2:C12").getValues()
    },
    {
      search: sheet2.getRange("E2:E8").getValues(),
      left: sheet2.getRange("D2:D8").getValues(),
      right: sheet2.getRange("F2:F8").getValues()
    },
    {
      search: sheet2.getRange("H2:H10").getValues(),
      left: sheet2.getRange("G2:G10").getValues(),
      right: sheet2.getRange("I2:I10").getValues()
    },
    {
      search: sheet2.getRange("K2:K9").getValues(),
      left: sheet2.getRange("J2:J9").getValues(),
      right: sheet2.getRange("L2:L9").getValues()
    }
  ];

  var upperRanges = [
    {
      search: sheet2.getRange("B23:B32").getValues(),
      left: sheet2.getRange("A23:A32").getValues(),
      right: sheet2.getRange("C23:C32").getValues()
    },
    {
      search: sheet2.getRange("E23:E28").getValues(),
      left: sheet2.getRange("D23:D28").getValues(),
      right: sheet2.getRange("F23:F28").getValues()
    },
    {
      search: sheet2.getRange("H23:H30").getValues(),
      left: sheet2.getRange("G23:G30").getValues(),
      right: sheet2.getRange("I23:I30").getValues()
    },
    {
      search: sheet2.getRange("K23:K29").getValues(),
      left: sheet2.getRange("J23:J29").getValues(),
      right: sheet2.getRange("L23:L29").getValues()
    }
  ];

  // Initialize the result variable
  var result = "";

  if (level <= 9){
    // Loop through each range group
    for (var i = 0; i < ranges.length; i++) {
      var searchRange = ranges[i].search;
      var leftRange = ranges[i].left;
      var rightRange = ranges[i].right;

    // Debugging: Check range dimensions
    Logger.log("Search length: " + searchRange.length + ", Left length: " + leftRange.length + ", Right length: " + rightRange.length);

    // Loop through each cell in the current search range
    for (var j = 0; j < searchRange.length; j++) {
      var currentSearchValue = searchRange[j][0].toString().trim().toLowerCase();
      var currentLeftValue = leftRange[j][0].toString().trim().toLowerCase();

      // Debugging: Log current values
      Logger.log("Current search value: " + currentSearchValue);
      Logger.log("Current left value: " + currentLeftValue);
      Logger.log("Right value: " + (rightRange[j][0] || "Undefined"));

      // Check if both A1 and B1 match the criteria (case-insensitive)
      if (currentSearchValue === valueA1 && currentLeftValue === valueB1) {
        result = rightRange[j][0]; // Get the value to the right of the matched cell
        break;
      }
    }

    if (result !== "") break; // Stop searching through other ranges if a match is found
  }
  } else if(level > 9) {
    // Loop through each range group
    for (var i = 0; i < upperRanges.length; i++) {
      var searchRange = upperRanges[i].search;
      var leftRange = upperRanges[i].left;
      var rightRange = upperRanges[i].right;

    // Debugging: Check range dimensions
    Logger.log("Search length: " + searchRange.length + ", Left length: " + leftRange.length + ", Right length: " + rightRange.length);

    // Loop through each cell in the current search range
    for (var j = 0; j < searchRange.length; j++) {
      var currentSearchValue = searchRange[j][0].toString().trim().toLowerCase();
      var currentLeftValue = leftRange[j][0].toString().trim().toLowerCase();

      // Debugging: Log current values
      Logger.log("Current search value: " + currentSearchValue);
      Logger.log("Current left value: " + currentLeftValue);
      Logger.log("Right value: " + (rightRange[j][0] || "Undefined"));

      // Check if both A1 and B1 match the criteria (case-insensitive)
      if (currentSearchValue === valueA1 && currentLeftValue === valueB1) {
        result = rightRange[j][0]; // Get the value to the right of the matched cell
        break;
      }
    }

    if (result !== "") break; // Stop searching through other ranges if a match is found
  }
  }

  // Clear the target cells before appending
  sheet1.getRange("I2").clearContent();
  sheet1.getRange("I3").clearContent();
  sheet1.getRange("M2").clearContent();

  var targetCell = sheet1.getRange("I2"); // Replace with your desired target cell
  var existingContent = targetCell.getValue();

  // Append the result on a new line if there's existing content
  if (existingContent) {
    targetCell.setValue(existingContent + "\n" + result);
  } else {
    targetCell.setValue(result);
  }

  var targetCell2 = sheet1.getRange("I3"); // Replace with your desired target cell
  var existingContent2 = targetCell2.getValue();

  // Append the result on a new line if there's existing content
  if (existingContent2) {
    targetCell2.setValue(existingContent2 + "\n" + valueA1);
  } else {
    targetCell2.setValue(valueA1);
  }
  // Print the result to a specific cell in the first sheet
  /*sheet1.getRange("I2").setValue(result); // Replace C1 with your desired result cell
  sheet1.getRange("I3").setValue(valueA1); // Replace C1 with your desired result cell*/
}

function matchAndPrintFromTwoSheets2() {
  // Get the spreadsheet and sheets
  /*var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet1 = spreadsheet.getSheetByName("Lesson Plan AI Prompt"); // Replace with the actual name of your first sheet
  var sheet2 = spreadsheet.getSheetByName("Skills Times"); // Replace with the actual name of your second sheet
  var sheet3 = spreadsheet.getSheetByName("Lesson Planning");*/

  if (!sheet1 || !sheet2 || !sheet3) {
    throw new Error("One or both sheets are not found. Ensure sheet names are correct.");
  }

  // Get values from the first sheet
  var valueA1 = sheet1.getRange("D3").getValue().toString().trim().toLowerCase();
  var valueB1 = sheet1.getRange("C3").getValue().toString().trim().toLowerCase();
  var level = sheet3.getRange("B3").getValue().toString().trim().toLowerCase();

  // Define 4 search ranges and their corresponding left and right ranges
  var ranges = [
    {
      search: sheet2.getRange("B2:B12").getValues(),
      left: sheet2.getRange("A2:A12").getValues(),
      right: sheet2.getRange("C2:C12").getValues()
    },
    {
      search: sheet2.getRange("E2:E8").getValues(),
      left: sheet2.getRange("D2:D8").getValues(),
      right: sheet2.getRange("F2:F8").getValues()
    },
    {
      search: sheet2.getRange("H2:H10").getValues(),
      left: sheet2.getRange("G2:G10").getValues(),
      right: sheet2.getRange("I2:I10").getValues()
    },
    {
      search: sheet2.getRange("K2:K9").getValues(),
      left: sheet2.getRange("J2:J9").getValues(),
      right: sheet2.getRange("L2:L9").getValues()
    }
  ];

  var upperRanges = [
    {
      search: sheet2.getRange("B23:B32").getValues(),
      left: sheet2.getRange("A23:A32").getValues(),
      right: sheet2.getRange("C23:C32").getValues()
    },
    {
      search: sheet2.getRange("E23:E28").getValues(),
      left: sheet2.getRange("D23:D28").getValues(),
      right: sheet2.getRange("F23:F28").getValues()
    },
    {
      search: sheet2.getRange("H23:H30").getValues(),
      left: sheet2.getRange("G23:G30").getValues(),
      right: sheet2.getRange("I23:I30").getValues()
    },
    {
      search: sheet2.getRange("K23:K29").getValues(),
      left: sheet2.getRange("J23:J29").getValues(),
      right: sheet2.getRange("L23:L29").getValues()
    }
  ];

  // Initialize the result variable
  var result = "";

  if (level <= 9){
    // Loop through each range group
    for (var i = 0; i < ranges.length; i++) {
      var searchRange = ranges[i].search;
      var leftRange = ranges[i].left;
      var rightRange = ranges[i].right;

    // Debugging: Check range dimensions
    Logger.log("Search length: " + searchRange.length + ", Left length: " + leftRange.length + ", Right length: " + rightRange.length);

    // Loop through each cell in the current search range
    for (var j = 0; j < searchRange.length; j++) {
      var currentSearchValue = searchRange[j][0].toString().trim().toLowerCase();
      var currentLeftValue = leftRange[j][0].toString().trim().toLowerCase();

      // Debugging: Log current values
      Logger.log("Current search value: " + currentSearchValue);
      Logger.log("Current left value: " + currentLeftValue);
      Logger.log("Right value: " + (rightRange[j][0] || "Undefined"));

      // Check if both A1 and B1 match the criteria (case-insensitive)
      if (currentSearchValue === valueA1 && currentLeftValue === valueB1) {
        result = rightRange[j][0]; // Get the value to the right of the matched cell
        break;
      }
    }

    if (result !== "") break; // Stop searching through other ranges if a match is found
  }
  } else if(level > 9) {
    // Loop through each range group
    for (var i = 0; i < upperRanges.length; i++) {
      var searchRange = upperRanges[i].search;
      var leftRange = upperRanges[i].left;
      var rightRange = upperRanges[i].right;

    // Debugging: Check range dimensions
    Logger.log("Search length: " + searchRange.length + ", Left length: " + leftRange.length + ", Right length: " + rightRange.length);

    // Loop through each cell in the current search range
    for (var j = 0; j < searchRange.length; j++) {
      var currentSearchValue = searchRange[j][0].toString().trim().toLowerCase();
      var currentLeftValue = leftRange[j][0].toString().trim().toLowerCase();

      // Debugging: Log current values
      Logger.log("Current search value: " + currentSearchValue);
      Logger.log("Current left value: " + currentLeftValue);
      Logger.log("Right value: " + (rightRange[j][0] || "Undefined"));

      // Check if both A1 and B1 match the criteria (case-insensitive)
      if (currentSearchValue === valueA1 && currentLeftValue === valueB1) {
        result = rightRange[j][0]; // Get the value to the right of the matched cell
        break;
      }
    }

    if (result !== "") break; // Stop searching through other ranges if a match is found
  }
  }

  var targetCell = sheet1.getRange("I2"); // Replace with your desired target cell
  var existingContent = targetCell.getValue();

  // Append the result on a new line if there's existing content
  if (existingContent) {
    targetCell.setValue(existingContent + "\n" + result);
  } else {
    targetCell.setValue(result);
  }

  var targetCell2 = sheet1.getRange("I3"); // Replace with your desired target cell
  var existingContent2 = targetCell2.getValue();

  // Append the result on a new line if there's existing content
  if (existingContent2) {
    targetCell2.setValue(existingContent2 + "\n" + valueA1);
  } else {
    targetCell2.setValue(valueA1);
  }
  // Print the result to a specific cell in the first sheet
  /*sheet1.getRange("I2").setValue(result); // Replace C1 with your desired result cell
  sheet1.getRange("I3").setValue(valueA1); // Replace C1 with your desired result cell*/
}

function matchAndPrintFromTwoSheets3() {
  // Get the spreadsheet and sheets
  /*var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet1 = spreadsheet.getSheetByName("Lesson Plan AI Prompt"); // Replace with the actual name of your first sheet
  var sheet2 = spreadsheet.getSheetByName("Skills Times"); // Replace with the actual name of your second sheet
  var sheet3 = spreadsheet.getSheetByName("Lesson Planning");*/

  if (!sheet1 || !sheet2 || !sheet3) {
    throw new Error("One or both sheets are not found. Ensure sheet names are correct.");
  }

  // Get values from the first sheet
  var valueA1 = sheet1.getRange("D4").getValue().toString().trim().toLowerCase();
  var valueB1 = sheet1.getRange("C4").getValue().toString().trim().toLowerCase();
  var level = sheet3.getRange("B3").getValue().toString().trim().toLowerCase();

  // Define 4 search ranges and their corresponding left and right ranges
  var ranges = [
    {
      search: sheet2.getRange("B2:B12").getValues(),
      left: sheet2.getRange("A2:A12").getValues(),
      right: sheet2.getRange("C2:C12").getValues()
    },
    {
      search: sheet2.getRange("E2:E8").getValues(),
      left: sheet2.getRange("D2:D8").getValues(),
      right: sheet2.getRange("F2:F8").getValues()
    },
    {
      search: sheet2.getRange("H2:H10").getValues(),
      left: sheet2.getRange("G2:G10").getValues(),
      right: sheet2.getRange("I2:I10").getValues()
    },
    {
      search: sheet2.getRange("K2:K9").getValues(),
      left: sheet2.getRange("J2:J9").getValues(),
      right: sheet2.getRange("L2:L9").getValues()
    }
  ];

  var upperRanges = [
    {
      search: sheet2.getRange("B23:B32").getValues(),
      left: sheet2.getRange("A23:A32").getValues(),
      right: sheet2.getRange("C23:C32").getValues()
    },
    {
      search: sheet2.getRange("E23:E28").getValues(),
      left: sheet2.getRange("D23:D28").getValues(),
      right: sheet2.getRange("F23:F28").getValues()
    },
    {
      search: sheet2.getRange("H23:H30").getValues(),
      left: sheet2.getRange("G23:G30").getValues(),
      right: sheet2.getRange("I23:I30").getValues()
    },
    {
      search: sheet2.getRange("K23:K29").getValues(),
      left: sheet2.getRange("J23:J29").getValues(),
      right: sheet2.getRange("L23:L29").getValues()
    }
  ];

  // Initialize the result variable
  var result = "";

  if (level <= 9){
    // Loop through each range group
    for (var i = 0; i < ranges.length; i++) {
      var searchRange = ranges[i].search;
      var leftRange = ranges[i].left;
      var rightRange = ranges[i].right;

    // Debugging: Check range dimensions
    Logger.log("Search length: " + searchRange.length + ", Left length: " + leftRange.length + ", Right length: " + rightRange.length);

    // Loop through each cell in the current search range
    for (var j = 0; j < searchRange.length; j++) {
      var currentSearchValue = searchRange[j][0].toString().trim().toLowerCase();
      var currentLeftValue = leftRange[j][0].toString().trim().toLowerCase();

      // Debugging: Log current values
      Logger.log("Current search value: " + currentSearchValue);
      Logger.log("Current left value: " + currentLeftValue);
      Logger.log("Right value: " + (rightRange[j][0] || "Undefined"));

      // Check if both A1 and B1 match the criteria (case-insensitive)
      if (currentSearchValue === valueA1 && currentLeftValue === valueB1) {
        result = rightRange[j][0]; // Get the value to the right of the matched cell
        break;
      }
    }

    if (result !== "") break; // Stop searching through other ranges if a match is found
  }
  } else if(level > 9) {
    // Loop through each range group
    for (var i = 0; i < upperRanges.length; i++) {
      var searchRange = upperRanges[i].search;
      var leftRange = upperRanges[i].left;
      var rightRange = upperRanges[i].right;

    // Debugging: Check range dimensions
    Logger.log("Search length: " + searchRange.length + ", Left length: " + leftRange.length + ", Right length: " + rightRange.length);

    // Loop through each cell in the current search range
    for (var j = 0; j < searchRange.length; j++) {
      var currentSearchValue = searchRange[j][0].toString().trim().toLowerCase();
      var currentLeftValue = leftRange[j][0].toString().trim().toLowerCase();

      // Debugging: Log current values
      Logger.log("Current search value: " + currentSearchValue);
      Logger.log("Current left value: " + currentLeftValue);
      Logger.log("Right value: " + (rightRange[j][0] || "Undefined"));

      // Check if both A1 and B1 match the criteria (case-insensitive)
      if (currentSearchValue === valueA1 && currentLeftValue === valueB1) {
        result = rightRange[j][0]; // Get the value to the right of the matched cell
        break;
      }
    }

    if (result !== "") break; // Stop searching through other ranges if a match is found
  }
  }

  var targetCell = sheet1.getRange("I2"); // Replace with your desired target cell
  var existingContent = targetCell.getValue();

  // Append the result on a new line if there's existing content
  if (existingContent) {
    targetCell.setValue(existingContent + "\n" + result);
  } else {
    targetCell.setValue(result);
  }

  var targetCell2 = sheet1.getRange("I3"); // Replace with your desired target cell
  var existingContent2 = targetCell2.getValue();

  // Append the result on a new line if there's existing content
  if (existingContent2) {
    targetCell2.setValue(existingContent2 + "\n" + valueA1);
  } else {
    targetCell2.setValue(valueA1);
  }
  // Print the result to a specific cell in the first sheet
  /*sheet1.getRange("I2").setValue(result); // Replace C1 with your desired result cell
  sheet1.getRange("I3").setValue(valueA1); // Replace C1 with your desired result cell*/
}

function matchAndPrintFromTwoSheets4() {
  // Get the spreadsheet and sheets
  /*var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet1 = spreadsheet.getSheetByName("Lesson Plan AI Prompt"); // Replace with the actual name of your first sheet
  var sheet2 = spreadsheet.getSheetByName("Skills Times"); // Replace with the actual name of your second sheet
  var sheet3 = spreadsheet.getSheetByName("Lesson Planning");*/

  if (!sheet1 || !sheet2 || !sheet3) {
    throw new Error("One or both sheets are not found. Ensure sheet names are correct.");
  }

  // Get values from the first sheet
  var valueA1 = sheet1.getRange("D5").getValue().toString().trim().toLowerCase();
  var valueB1 = sheet1.getRange("C5").getValue().toString().trim().toLowerCase();
  var level = sheet3.getRange("B3").getValue().toString().trim().toLowerCase();

  // Define 4 search ranges and their corresponding left and right ranges
  var ranges = [
    {
      search: sheet2.getRange("B2:B12").getValues(),
      left: sheet2.getRange("A2:A12").getValues(),
      right: sheet2.getRange("C2:C12").getValues()
    },
    {
      search: sheet2.getRange("E2:E8").getValues(),
      left: sheet2.getRange("D2:D8").getValues(),
      right: sheet2.getRange("F2:F8").getValues()
    },
    {
      search: sheet2.getRange("H2:H10").getValues(),
      left: sheet2.getRange("G2:G10").getValues(),
      right: sheet2.getRange("I2:I10").getValues()
    },
    {
      search: sheet2.getRange("K2:K9").getValues(),
      left: sheet2.getRange("J2:J9").getValues(),
      right: sheet2.getRange("L2:L9").getValues()
    }
  ];

  var upperRanges = [
    {
      search: sheet2.getRange("B23:B32").getValues(),
      left: sheet2.getRange("A23:A32").getValues(),
      right: sheet2.getRange("C23:C32").getValues()
    },
    {
      search: sheet2.getRange("E23:E28").getValues(),
      left: sheet2.getRange("D23:D28").getValues(),
      right: sheet2.getRange("F23:F28").getValues()
    },
    {
      search: sheet2.getRange("H23:H30").getValues(),
      left: sheet2.getRange("G23:G30").getValues(),
      right: sheet2.getRange("I23:I30").getValues()
    },
    {
      search: sheet2.getRange("K23:K29").getValues(),
      left: sheet2.getRange("J23:J29").getValues(),
      right: sheet2.getRange("L23:L29").getValues()
    }
  ];

  // Initialize the result variable
  var result = "";

  if (level <= 9){
    // Loop through each range group
    for (var i = 0; i < ranges.length; i++) {
      var searchRange = ranges[i].search;
      var leftRange = ranges[i].left;
      var rightRange = ranges[i].right;

    // Debugging: Check range dimensions
    Logger.log("Search length: " + searchRange.length + ", Left length: " + leftRange.length + ", Right length: " + rightRange.length);

    // Loop through each cell in the current search range
    for (var j = 0; j < searchRange.length; j++) {
      var currentSearchValue = searchRange[j][0].toString().trim().toLowerCase();
      var currentLeftValue = leftRange[j][0].toString().trim().toLowerCase();

      // Debugging: Log current values
      Logger.log("Current search value: " + currentSearchValue);
      Logger.log("Current left value: " + currentLeftValue);
      Logger.log("Right value: " + (rightRange[j][0] || "Undefined"));

      // Check if both A1 and B1 match the criteria (case-insensitive)
      if (currentSearchValue === valueA1 && currentLeftValue === valueB1) {
        result = rightRange[j][0]; // Get the value to the right of the matched cell
        break;
      }
    }

    if (result !== "") break; // Stop searching through other ranges if a match is found
  }
  } else if(level > 9) {
    // Loop through each range group
    for (var i = 0; i < upperRanges.length; i++) {
      var searchRange = upperRanges[i].search;
      var leftRange = upperRanges[i].left;
      var rightRange = upperRanges[i].right;

    // Debugging: Check range dimensions
    Logger.log("Search length: " + searchRange.length + ", Left length: " + leftRange.length + ", Right length: " + rightRange.length);

    // Loop through each cell in the current search range
    for (var j = 0; j < searchRange.length; j++) {
      var currentSearchValue = searchRange[j][0].toString().trim().toLowerCase();
      var currentLeftValue = leftRange[j][0].toString().trim().toLowerCase();

      // Debugging: Log current values
      Logger.log("Current search value: " + currentSearchValue);
      Logger.log("Current left value: " + currentLeftValue);
      Logger.log("Right value: " + (rightRange[j][0] || "Undefined"));

      // Check if both A1 and B1 match the criteria (case-insensitive)
      if (currentSearchValue === valueA1 && currentLeftValue === valueB1) {
        result = rightRange[j][0]; // Get the value to the right of the matched cell
        break;
      }
    }

    if (result !== "") break; // Stop searching through other ranges if a match is found
  }
  }

  var targetCell = sheet1.getRange("I2"); // Replace with your desired target cell
  var existingContent = targetCell.getValue();

  // Append the result on a new line if there's existing content
  if (existingContent) {
    targetCell.setValue(existingContent + "\n" + result);
  } else {
    targetCell.setValue(result);
  }

  var targetCell2 = sheet1.getRange("I3"); // Replace with your desired target cell
  var existingContent2 = targetCell2.getValue();

  // Append the result on a new line if there's existing content
  if (existingContent2) {
    targetCell2.setValue(existingContent2 + "\n" + valueA1);
  } else {
    targetCell2.setValue(valueA1);
  }
  // Print the result to a specific cell in the first sheet
  /*sheet1.getRange("I2").setValue(result); // Replace C1 with your desired result cell
  sheet1.getRange("I3").setValue(valueA1); // Replace C1 with your desired result cell*/
}


function matchAndPrintFromTwoSheets5() {
  // Get the spreadsheet and sheets
  /*var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet1 = spreadsheet.getSheetByName("Lesson Plan AI Prompt"); // Replace with the actual name of your first sheet
  var sheet2 = spreadsheet.getSheetByName("Skills Times"); // Replace with the actual name of your second sheet
  var sheet3 = spreadsheet.getSheetByName("Lesson Planning");*/

  if (!sheet1 || !sheet2 || !sheet3) {
    throw new Error("One or both sheets are not found. Ensure sheet names are correct.");
  }

  // Get values from the first sheet
  var valueA1 = sheet1.getRange("D6").getValue().toString().trim().toLowerCase();
  var valueB1 = sheet1.getRange("C6").getValue().toString().trim().toLowerCase();
  var level = sheet3.getRange("B3").getValue().toString().trim().toLowerCase();

  // Define 4 search ranges and their corresponding left and right ranges
  var ranges = [
    {
      search: sheet2.getRange("B2:B12").getValues(),
      left: sheet2.getRange("A2:A12").getValues(),
      right: sheet2.getRange("C2:C12").getValues()
    },
    {
      search: sheet2.getRange("E2:E8").getValues(),
      left: sheet2.getRange("D2:D8").getValues(),
      right: sheet2.getRange("F2:F8").getValues()
    },
    {
      search: sheet2.getRange("H2:H10").getValues(),
      left: sheet2.getRange("G2:G10").getValues(),
      right: sheet2.getRange("I2:I10").getValues()
    },
    {
      search: sheet2.getRange("K2:K9").getValues(),
      left: sheet2.getRange("J2:J9").getValues(),
      right: sheet2.getRange("L2:L9").getValues()
    }
  ];

  var upperRanges = [
    {
      search: sheet2.getRange("B23:B32").getValues(),
      left: sheet2.getRange("A23:A32").getValues(),
      right: sheet2.getRange("C23:C32").getValues()
    },
    {
      search: sheet2.getRange("E23:E28").getValues(),
      left: sheet2.getRange("D23:D28").getValues(),
      right: sheet2.getRange("F23:F28").getValues()
    },
    {
      search: sheet2.getRange("H23:H30").getValues(),
      left: sheet2.getRange("G23:G30").getValues(),
      right: sheet2.getRange("I23:I30").getValues()
    },
    {
      search: sheet2.getRange("K23:K29").getValues(),
      left: sheet2.getRange("J23:J29").getValues(),
      right: sheet2.getRange("L23:L29").getValues()
    }
  ];

  // Initialize the result variable
  var result = "";

  if (level <= 9){
    // Loop through each range group
    for (var i = 0; i < ranges.length; i++) {
      var searchRange = ranges[i].search;
      var leftRange = ranges[i].left;
      var rightRange = ranges[i].right;

    // Debugging: Check range dimensions
    Logger.log("Search length: " + searchRange.length + ", Left length: " + leftRange.length + ", Right length: " + rightRange.length);

    // Loop through each cell in the current search range
    for (var j = 0; j < searchRange.length; j++) {
      var currentSearchValue = searchRange[j][0].toString().trim().toLowerCase();
      var currentLeftValue = leftRange[j][0].toString().trim().toLowerCase();

      // Debugging: Log current values
      Logger.log("Current search value: " + currentSearchValue);
      Logger.log("Current left value: " + currentLeftValue);
      Logger.log("Right value: " + (rightRange[j][0] || "Undefined"));

      // Check if both A1 and B1 match the criteria (case-insensitive)
      if (currentSearchValue === valueA1 && currentLeftValue === valueB1) {
        result = rightRange[j][0]; // Get the value to the right of the matched cell
        break;
      }
    }

    if (result !== "") break; // Stop searching through other ranges if a match is found
  }
  } else if(level > 9) {
    // Loop through each range group
    for (var i = 0; i < upperRanges.length; i++) {
      var searchRange = upperRanges[i].search;
      var leftRange = upperRanges[i].left;
      var rightRange = upperRanges[i].right;

    // Debugging: Check range dimensions
    Logger.log("Search length: " + searchRange.length + ", Left length: " + leftRange.length + ", Right length: " + rightRange.length);

    // Loop through each cell in the current search range
    for (var j = 0; j < searchRange.length; j++) {
      var currentSearchValue = searchRange[j][0].toString().trim().toLowerCase();
      var currentLeftValue = leftRange[j][0].toString().trim().toLowerCase();

      // Debugging: Log current values
      Logger.log("Current search value: " + currentSearchValue);
      Logger.log("Current left value: " + currentLeftValue);
      Logger.log("Right value: " + (rightRange[j][0] || "Undefined"));

      // Check if both A1 and B1 match the criteria (case-insensitive)
      if (currentSearchValue === valueA1 && currentLeftValue === valueB1) {
        result = rightRange[j][0]; // Get the value to the right of the matched cell
        break;
      }
    }

    if (result !== "") break; // Stop searching through other ranges if a match is found
  }
  }

  var targetCell = sheet1.getRange("I2"); // Replace with your desired target cell
  var existingContent = targetCell.getValue();

  // Append the result on a new line if there's existing content
  if (existingContent) {
    targetCell.setValue(existingContent + "\n" + result);
  } else {
    targetCell.setValue(result);
  }

  var targetCell2 = sheet1.getRange("I3"); // Replace with your desired target cell
  var existingContent2 = targetCell2.getValue();

  // Append the result on a new line if there's existing content
  if (existingContent2) {
    targetCell2.setValue(existingContent2 + "\n" + valueA1);
  } else {
    targetCell2.setValue(valueA1);
  }
  // Print the result to a specific cell in the first sheet
  /*sheet1.getRange("I2").setValue(result); // Replace C1 with your desired result cell
  sheet1.getRange("I3").setValue(valueA1); // Replace C1 with your desired result cell*/
}


function matchAndPrintFromTwoSheets6() {
  // Get the spreadsheet and sheets
  /*var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet1 = spreadsheet.getSheetByName("Lesson Plan AI Prompt"); // Replace with the actual name of your first sheet
  var sheet2 = spreadsheet.getSheetByName("Skills Times"); // Replace with the actual name of your second sheet
  var sheet3 = spreadsheet.getSheetByName("Lesson Planning");*/

  if (!sheet1 || !sheet2 || !sheet3) {
    throw new Error("One or both sheets are not found. Ensure sheet names are correct.");
  }

  // Get values from the first sheet
  var valueA1 = sheet1.getRange("D7").getValue().toString().trim().toLowerCase();
  var valueB1 = sheet1.getRange("C7").getValue().toString().trim().toLowerCase();
  var level = sheet3.getRange("B3").getValue().toString().trim().toLowerCase();

  // Define 4 search ranges and their corresponding left and right ranges
  var ranges = [
    {
      search: sheet2.getRange("B2:B12").getValues(),
      left: sheet2.getRange("A2:A12").getValues(),
      right: sheet2.getRange("C2:C12").getValues()
    },
    {
      search: sheet2.getRange("E2:E8").getValues(),
      left: sheet2.getRange("D2:D8").getValues(),
      right: sheet2.getRange("F2:F8").getValues()
    },
    {
      search: sheet2.getRange("H2:H10").getValues(),
      left: sheet2.getRange("G2:G10").getValues(),
      right: sheet2.getRange("I2:I10").getValues()
    },
    {
      search: sheet2.getRange("K2:K9").getValues(),
      left: sheet2.getRange("J2:J9").getValues(),
      right: sheet2.getRange("L2:L9").getValues()
    }
  ];

  var upperRanges = [
    {
      search: sheet2.getRange("B23:B32").getValues(),
      left: sheet2.getRange("A23:A32").getValues(),
      right: sheet2.getRange("C23:C32").getValues()
    },
    {
      search: sheet2.getRange("E23:E28").getValues(),
      left: sheet2.getRange("D23:D28").getValues(),
      right: sheet2.getRange("F23:F28").getValues()
    },
    {
      search: sheet2.getRange("H23:H30").getValues(),
      left: sheet2.getRange("G23:G30").getValues(),
      right: sheet2.getRange("I23:I30").getValues()
    },
    {
      search: sheet2.getRange("K23:K29").getValues(),
      left: sheet2.getRange("J23:J29").getValues(),
      right: sheet2.getRange("L23:L29").getValues()
    }
  ];

  // Initialize the result variable
  var result = "";

  if (level <= 9){
    // Loop through each range group
    for (var i = 0; i < ranges.length; i++) {
      var searchRange = ranges[i].search;
      var leftRange = ranges[i].left;
      var rightRange = ranges[i].right;

    // Debugging: Check range dimensions
    Logger.log("Search length: " + searchRange.length + ", Left length: " + leftRange.length + ", Right length: " + rightRange.length);

    // Loop through each cell in the current search range
    for (var j = 0; j < searchRange.length; j++) {
      var currentSearchValue = searchRange[j][0].toString().trim().toLowerCase();
      var currentLeftValue = leftRange[j][0].toString().trim().toLowerCase();

      // Debugging: Log current values
      Logger.log("Current search value: " + currentSearchValue);
      Logger.log("Current left value: " + currentLeftValue);
      Logger.log("Right value: " + (rightRange[j][0] || "Undefined"));

      // Check if both A1 and B1 match the criteria (case-insensitive)
      if (currentSearchValue === valueA1 && currentLeftValue === valueB1) {
        result = rightRange[j][0]; // Get the value to the right of the matched cell
        break;
      }
    }

    if (result !== "") break; // Stop searching through other ranges if a match is found
  }
  } else if(level > 9) {
    // Loop through each range group
    for (var i = 0; i < upperRanges.length; i++) {
      var searchRange = upperRanges[i].search;
      var leftRange = upperRanges[i].left;
      var rightRange = upperRanges[i].right;

    // Debugging: Check range dimensions
    Logger.log("Search length: " + searchRange.length + ", Left length: " + leftRange.length + ", Right length: " + rightRange.length);

    // Loop through each cell in the current search range
    for (var j = 0; j < searchRange.length; j++) {
      var currentSearchValue = searchRange[j][0].toString().trim().toLowerCase();
      var currentLeftValue = leftRange[j][0].toString().trim().toLowerCase();

      // Debugging: Log current values
      Logger.log("Current search value: " + currentSearchValue);
      Logger.log("Current left value: " + currentLeftValue);
      Logger.log("Right value: " + (rightRange[j][0] || "Undefined"));

      // Check if both A1 and B1 match the criteria (case-insensitive)
      if (currentSearchValue === valueA1 && currentLeftValue === valueB1) {
        result = rightRange[j][0]; // Get the value to the right of the matched cell
        break;
      }
    }

    if (result !== "") break; // Stop searching through other ranges if a match is found
  }
  }

  var targetCell = sheet1.getRange("I2"); // Replace with your desired target cell
  var existingContent = targetCell.getValue();

  // Append the result on a new line if there's existing content
  if (existingContent) {
    targetCell.setValue(existingContent + "\n" + result);
  } else {
    targetCell.setValue(result);
  }

  var targetCell2 = sheet1.getRange("I3"); // Replace with your desired target cell
  var existingContent2 = targetCell2.getValue();

  // Append the result on a new line if there's existing content
  if (existingContent2) {
    targetCell2.setValue(existingContent2 + "\n" + valueA1);
  } else {
    targetCell2.setValue(valueA1);
  }
  // Print the result to a specific cell in the first sheet
  /*sheet1.getRange("I2").setValue(result); // Replace C1 with your desired result cell
  sheet1.getRange("I3").setValue(valueA1); // Replace C1 with your desired result cell*/
}



