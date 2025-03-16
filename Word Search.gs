/**
 * This function searches for multiple words from multiple cells in a target cell (case-insensitive).
 * It checks each word or phrase in the source cells and searches it in the target cell.
 * 
 * @param {Array} wordsArray An array of words/phrases to search for.
 * @param {string} targetText The text where we will search for the words/phrases.
 * @return {string} Returns a message listing all words that were found or not found.
 */
function searchMultipleWords(wordsArray, targetText) {
  var result = "";  // Initialize an empty result string
  
  // Convert the target text to lowercase to make the search case-insensitive
  var lowerTargetText = targetText.toLowerCase();

  // Loop through each word or phrase in the wordsArray
  wordsArray.forEach(function(word) {
    // Ensure word is a string and convert it to lowercase
    if (typeof word === 'string') {
      var lowerWord = word.toLowerCase();
      
      // Check if the word/phrase is found in the target text
      if (lowerTargetText.indexOf(lowerWord) !== -1) {
        result += '' + word + '\n';  // Add the found word to the result
      } else {
        result += '';  // Add the not found word to the result
      }
    }
  });
  
  return result;  // Return the result string
}

var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();


/**
 * This function is triggered from the custom menu to run the search for multiple words.
 */
function runSearch() {
  //var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Get the target text from cell B1 (you can change this to any target cell)
  var targetText = sheet.getRange('AI Prompt!A4').getValue();
  
  // Get words/phrases from 4 different ranges:
  var wordsRange1 = sheet.getRange('Resources!A4:A32').getValues();
  var wordsRange2 = sheet.getRange('Resources!B4:B33').getValues();
  var wordsRange3 = sheet.getRange('Resources!C4:C42').getValues();
  var wordsRange4 = sheet.getRange('Resources!D4:D24').getValues();
  
  // Flatten each array of words (to avoid 2D arrays)
  var wordsArray1 = wordsRange1.map(function(row) { return row[0]; });
  var wordsArray2 = wordsRange2.map(function(row) { return row[0]; });
  var wordsArray3 = wordsRange3.map(function(row) { return row[0]; });
  var wordsArray4 = wordsRange4.map(function(row) { return row[0]; });
  
  // Call the function to search for words in the target text for each set of words
  var result1 = searchMultipleWords(wordsArray1, targetText);
  var result2 = searchMultipleWords(wordsArray2, targetText);
  var result3 = searchMultipleWords(wordsArray3, targetText);
  var result4 = searchMultipleWords(wordsArray4, targetText);
  
  // Output the results in cells C1, C2, C3, and C4
  /*sheet.getRange('AI Prompt!B8').setValue(result1); 
  sheet.getRange('AI Prompt!B9').setValue(result2);
  sheet.getRange('AI Prompt!B10').setValue(result3);
  sheet.getRange('AI Prompt!B11').setValue(result4);*/

  // Combine all results into one string
  var combinedResults = "";
  if (result1) combinedResults += result1 + "\n";
  if (result2) combinedResults += result2 + "\n";
  if (result3) combinedResults += result3 + "\n";
  if (result4) combinedResults += result4 + "\n";
  
  // Output the combined results into a single cell, e.g., AI Prompt!B8
  sheet.getRange('AI Prompt!E4').setValue(combinedResults);
}

function runSearch2() {
  //var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Get the target text from cell B1 (you can change this to any target cell)
  var targetText = sheet.getRange('AI Prompt!K4').getValue();
  
  // Get words/phrases from 4 different ranges:
  var wordsRange1 = sheet.getRange('Resources!A4:A32').getValues();
  var wordsRange2 = sheet.getRange('Resources!B4:B33').getValues();
  var wordsRange3 = sheet.getRange('Resources!C4:C42').getValues();
  var wordsRange4 = sheet.getRange('Resources!D4:D24').getValues();
  
  // Flatten each array of words (to avoid 2D arrays)
  var wordsArray1 = wordsRange1.map(function(row) { return row[0]; });
  var wordsArray2 = wordsRange2.map(function(row) { return row[0]; });
  var wordsArray3 = wordsRange3.map(function(row) { return row[0]; });
  var wordsArray4 = wordsRange4.map(function(row) { return row[0]; });
  
  // Call the function to search for words in the target text for each set of words
  var result1 = searchMultipleWords(wordsArray1, targetText);
  var result2 = searchMultipleWords(wordsArray2, targetText);
  var result3 = searchMultipleWords(wordsArray3, targetText);
  var result4 = searchMultipleWords(wordsArray4, targetText);
  
  // Output the results in cells C1, C2, C3, and C4
  /*sheet.getRange('AI Prompt!B8').setValue(result1); 
  sheet.getRange('AI Prompt!B9').setValue(result2);
  sheet.getRange('AI Prompt!B10').setValue(result3);
  sheet.getRange('AI Prompt!B11').setValue(result4);*/

  // Combine all results into one string
  var combinedResults = "";
  if (result1) combinedResults += result1 + "\n";
  if (result2) combinedResults += result2 + "\n";
  if (result3) combinedResults += result3 + "\n";
  if (result4) combinedResults += result4 + "\n";
  
  // Output the combined results into a single cell, e.g., AI Prompt!B8
  sheet.getRange('AI Prompt!O4').setValue(combinedResults);
}

function runSearch3() {
  //var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Get the target text from cell B1 (you can change this to any target cell)
  var targetText = sheet.getRange('AI Prompt!U4').getValue();
  
  // Get words/phrases from 4 different ranges:
  var wordsRange1 = sheet.getRange('Resources!A4:A32').getValues();
  var wordsRange2 = sheet.getRange('Resources!B4:B33').getValues();
  var wordsRange3 = sheet.getRange('Resources!C4:C42').getValues();
  var wordsRange4 = sheet.getRange('Resources!D4:D24').getValues();
  
  // Flatten each array of words (to avoid 2D arrays)
  var wordsArray1 = wordsRange1.map(function(row) { return row[0]; });
  var wordsArray2 = wordsRange2.map(function(row) { return row[0]; });
  var wordsArray3 = wordsRange3.map(function(row) { return row[0]; });
  var wordsArray4 = wordsRange4.map(function(row) { return row[0]; });
  
  // Call the function to search for words in the target text for each set of words
  var result1 = searchMultipleWords(wordsArray1, targetText);
  var result2 = searchMultipleWords(wordsArray2, targetText);
  var result3 = searchMultipleWords(wordsArray3, targetText);
  var result4 = searchMultipleWords(wordsArray4, targetText);
  
  // Output the results in cells C1, C2, C3, and C4
  /*sheet.getRange('AI Prompt!B8').setValue(result1); 
  sheet.getRange('AI Prompt!B9').setValue(result2);
  sheet.getRange('AI Prompt!B10').setValue(result3);
  sheet.getRange('AI Prompt!B11').setValue(result4);*/

  // Combine all results into one string
  var combinedResults = "";
  if (result1) combinedResults += result1 + "\n";
  if (result2) combinedResults += result2 + "\n";
  if (result3) combinedResults += result3 + "\n";
  if (result4) combinedResults += result4 + "\n";
  
  // Output the combined results into a single cell, e.g., AI Prompt!B8
  sheet.getRange('AI Prompt!Y4').setValue(combinedResults);
}

function runSearch4() {
  //var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Get the target text from cell B1 (you can change this to any target cell)
  var targetText = sheet.getRange('AI Prompt!AE4').getValue();
  
  // Get words/phrases from 4 different ranges:
  var wordsRange1 = sheet.getRange('Resources!A4:A32').getValues();
  var wordsRange2 = sheet.getRange('Resources!B4:B33').getValues();
  var wordsRange3 = sheet.getRange('Resources!C4:C42').getValues();
  var wordsRange4 = sheet.getRange('Resources!D4:D24').getValues();
  
  // Flatten each array of words (to avoid 2D arrays)
  var wordsArray1 = wordsRange1.map(function(row) { return row[0]; });
  var wordsArray2 = wordsRange2.map(function(row) { return row[0]; });
  var wordsArray3 = wordsRange3.map(function(row) { return row[0]; });
  var wordsArray4 = wordsRange4.map(function(row) { return row[0]; });
  
  // Call the function to search for words in the target text for each set of words
  var result1 = searchMultipleWords(wordsArray1, targetText);
  var result2 = searchMultipleWords(wordsArray2, targetText);
  var result3 = searchMultipleWords(wordsArray3, targetText);
  var result4 = searchMultipleWords(wordsArray4, targetText);
  
  // Output the results in cells C1, C2, C3, and C4
  /*sheet.getRange('AI Prompt!B8').setValue(result1); 
  sheet.getRange('AI Prompt!B9').setValue(result2);
  sheet.getRange('AI Prompt!B10').setValue(result3);
  sheet.getRange('AI Prompt!B11').setValue(result4);*/

  // Combine all results into one string
  var combinedResults = "";
  if (result1) combinedResults += result1 + "\n";
  if (result2) combinedResults += result2 + "\n";
  if (result3) combinedResults += result3 + "\n";
  if (result4) combinedResults += result4 + "\n";
  
  // Output the combined results into a single cell, e.g., AI Prompt!B8
  sheet.getRange('AI Prompt!AI4').setValue(combinedResults);
}

function runSearch5() {
  //var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Get the target text from cell B1 (you can change this to any target cell)
  var targetText = sheet.getRange('AI Prompt!AO4').getValue();
  
  // Get words/phrases from 4 different ranges:
  var wordsRange1 = sheet.getRange('Resources!A4:A32').getValues();
  var wordsRange2 = sheet.getRange('Resources!B4:B33').getValues();
  var wordsRange3 = sheet.getRange('Resources!C4:C42').getValues();
  var wordsRange4 = sheet.getRange('Resources!D4:D24').getValues();
  
  // Flatten each array of words (to avoid 2D arrays)
  var wordsArray1 = wordsRange1.map(function(row) { return row[0]; });
  var wordsArray2 = wordsRange2.map(function(row) { return row[0]; });
  var wordsArray3 = wordsRange3.map(function(row) { return row[0]; });
  var wordsArray4 = wordsRange4.map(function(row) { return row[0]; });
  
  // Call the function to search for words in the target text for each set of words
  var result1 = searchMultipleWords(wordsArray1, targetText);
  var result2 = searchMultipleWords(wordsArray2, targetText);
  var result3 = searchMultipleWords(wordsArray3, targetText);
  var result4 = searchMultipleWords(wordsArray4, targetText);
  
  // Output the results in cells C1, C2, C3, and C4
  /*sheet.getRange('AI Prompt!B8').setValue(result1); 
  sheet.getRange('AI Prompt!B9').setValue(result2);
  sheet.getRange('AI Prompt!B10').setValue(result3);
  sheet.getRange('AI Prompt!B11').setValue(result4);*/

  // Combine all results into one string
  var combinedResults = "";
  if (result1) combinedResults += result1 + "\n";
  if (result2) combinedResults += result2 + "\n";
  if (result3) combinedResults += result3 + "\n";
  if (result4) combinedResults += result4 + "\n";
  
  // Output the combined results into a single cell, e.g., AI Prompt!B8
  sheet.getRange('AI Prompt!AS4').setValue(combinedResults);
}

function runSearch6() {
  //var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Get the target text from cell B1 (you can change this to any target cell)
  var targetText = sheet.getRange('AI Prompt!AY4').getValue();
  
  // Get words/phrases from 4 different ranges:
  var wordsRange1 = sheet.getRange('Resources!A4:A32').getValues();
  var wordsRange2 = sheet.getRange('Resources!B4:B33').getValues();
  var wordsRange3 = sheet.getRange('Resources!C4:C42').getValues();
  var wordsRange4 = sheet.getRange('Resources!D4:D24').getValues();
  
  // Flatten each array of words (to avoid 2D arrays)
  var wordsArray1 = wordsRange1.map(function(row) { return row[0]; });
  var wordsArray2 = wordsRange2.map(function(row) { return row[0]; });
  var wordsArray3 = wordsRange3.map(function(row) { return row[0]; });
  var wordsArray4 = wordsRange4.map(function(row) { return row[0]; });
  
  // Call the function to search for words in the target text for each set of words
  var result1 = searchMultipleWords(wordsArray1, targetText);
  var result2 = searchMultipleWords(wordsArray2, targetText);
  var result3 = searchMultipleWords(wordsArray3, targetText);
  var result4 = searchMultipleWords(wordsArray4, targetText);
  
  // Output the results in cells C1, C2, C3, and C4
  /*sheet.getRange('AI Prompt!B8').setValue(result1); 
  sheet.getRange('AI Prompt!B9').setValue(result2);
  sheet.getRange('AI Prompt!B10').setValue(result3);
  sheet.getRange('AI Prompt!B11').setValue(result4);*/

  // Combine all results into one string
  var combinedResults = "";
  if (result1) combinedResults += result1 + "\n";
  if (result2) combinedResults += result2 + "\n";
  if (result3) combinedResults += result3 + "\n";
  if (result4) combinedResults += result4 + "\n";
  
  // Output the combined results into a single cell, e.g., AI Prompt!B8
  sheet.getRange('AI Prompt!BC4').setValue(combinedResults);
}

function runSearchMATSCI() {
  //var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Get the target text from cell B1 (you can change this to any target cell)
  var targetText = sheet.getRange('AI Prompt!A4').getValue();
  
  // Get words/phrases from 4 different ranges:
  var wordsRange1 = sheet.getRange('Resources!A48:A61').getValues();
  var wordsRange2 = sheet.getRange('Resources!B48:B61').getValues();
  var wordsRange3 = sheet.getRange('Resources!C48:C66').getValues();
  var wordsRange4 = sheet.getRange('Resources!D48:D55').getValues();
  
  // Flatten each array of words (to avoid 2D arrays)
  var wordsArray1 = wordsRange1.map(function(row) { return row[0]; });
  var wordsArray2 = wordsRange2.map(function(row) { return row[0]; });
  var wordsArray3 = wordsRange3.map(function(row) { return row[0]; });
  var wordsArray4 = wordsRange4.map(function(row) { return row[0]; });
  
  // Call the function to search for words in the target text for each set of words
  var result1 = searchMultipleWords(wordsArray1, targetText);
  var result2 = searchMultipleWords(wordsArray2, targetText);
  var result3 = searchMultipleWords(wordsArray3, targetText);
  var result4 = searchMultipleWords(wordsArray4, targetText);
  
  // Output the results in cells C1, C2, C3, and C4
  /*sheet.getRange('AI Prompt!B8').setValue(result1); 
  sheet.getRange('AI Prompt!B9').setValue(result2);
  sheet.getRange('AI Prompt!B10').setValue(result3);
  sheet.getRange('AI Prompt!B11').setValue(result4);*/

  // Combine all results into one string
  var combinedResults = "";
  if (result1) combinedResults += result1 + "\n";
  if (result2) combinedResults += result2 + "\n";
  if (result3) combinedResults += result3 + "\n";
  if (result4) combinedResults += result4 + "\n";
  
  // Output the combined results into a single cell, e.g., AI Prompt!B8
  sheet.getRange('AI Prompt!E4').setValue(combinedResults);
}

function runSearchMATSCI2() {
  //var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Get the target text from cell B1 (you can change this to any target cell)
  var targetText = sheet.getRange('AI Prompt!K4').getValue();
  
  // Get words/phrases from 4 different ranges:
  var wordsRange1 = sheet.getRange('Resources!A48:A61').getValues();
  var wordsRange2 = sheet.getRange('Resources!B48:B61').getValues();
  var wordsRange3 = sheet.getRange('Resources!C48:C66').getValues();
  var wordsRange4 = sheet.getRange('Resources!D48:D55').getValues();
  
  // Flatten each array of words (to avoid 2D arrays)
  var wordsArray1 = wordsRange1.map(function(row) { return row[0]; });
  var wordsArray2 = wordsRange2.map(function(row) { return row[0]; });
  var wordsArray3 = wordsRange3.map(function(row) { return row[0]; });
  var wordsArray4 = wordsRange4.map(function(row) { return row[0]; });
  
  // Call the function to search for words in the target text for each set of words
  var result1 = searchMultipleWords(wordsArray1, targetText);
  var result2 = searchMultipleWords(wordsArray2, targetText);
  var result3 = searchMultipleWords(wordsArray3, targetText);
  var result4 = searchMultipleWords(wordsArray4, targetText);
  
  // Output the results in cells C1, C2, C3, and C4
  /*sheet.getRange('AI Prompt!B8').setValue(result1); 
  sheet.getRange('AI Prompt!B9').setValue(result2);
  sheet.getRange('AI Prompt!B10').setValue(result3);
  sheet.getRange('AI Prompt!B11').setValue(result4);*/

  // Combine all results into one string
  var combinedResults = "";
  if (result1) combinedResults += result1 + "\n";
  if (result2) combinedResults += result2 + "\n";
  if (result3) combinedResults += result3 + "\n";
  if (result4) combinedResults += result4 + "\n";
  
  // Output the combined results into a single cell, e.g., AI Prompt!B8
  sheet.getRange('AI Prompt!O4').setValue(combinedResults);
}

function runSearchMATSCI3() {
  //var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Get the target text from cell B1 (you can change this to any target cell)
  var targetText = sheet.getRange('AI Prompt!U4').getValue();
  
  // Get words/phrases from 4 different ranges:
  var wordsRange1 = sheet.getRange('Resources!A48:A61').getValues();
  var wordsRange2 = sheet.getRange('Resources!B48:B61').getValues();
  var wordsRange3 = sheet.getRange('Resources!C48:C66').getValues();
  var wordsRange4 = sheet.getRange('Resources!D48:D55').getValues();
  
  // Flatten each array of words (to avoid 2D arrays)
  var wordsArray1 = wordsRange1.map(function(row) { return row[0]; });
  var wordsArray2 = wordsRange2.map(function(row) { return row[0]; });
  var wordsArray3 = wordsRange3.map(function(row) { return row[0]; });
  var wordsArray4 = wordsRange4.map(function(row) { return row[0]; });
  
  // Call the function to search for words in the target text for each set of words
  var result1 = searchMultipleWords(wordsArray1, targetText);
  var result2 = searchMultipleWords(wordsArray2, targetText);
  var result3 = searchMultipleWords(wordsArray3, targetText);
  var result4 = searchMultipleWords(wordsArray4, targetText);
  
  // Output the results in cells C1, C2, C3, and C4
  /*sheet.getRange('AI Prompt!B8').setValue(result1); 
  sheet.getRange('AI Prompt!B9').setValue(result2);
  sheet.getRange('AI Prompt!B10').setValue(result3);
  sheet.getRange('AI Prompt!B11').setValue(result4);*/

  // Combine all results into one string
  var combinedResults = "";
  if (result1) combinedResults += result1 + "\n";
  if (result2) combinedResults += result2 + "\n";
  if (result3) combinedResults += result3 + "\n";
  if (result4) combinedResults += result4 + "\n";
  
  // Output the combined results into a single cell, e.g., AI Prompt!B8
  sheet.getRange('AI Prompt!Y4').setValue(combinedResults);
}

function runSearchMATSCI4() {
  //var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Get the target text from cell B1 (you can change this to any target cell)
  var targetText = sheet.getRange('AI Prompt!AE4').getValue();
  
  // Get words/phrases from 4 different ranges:
  var wordsRange1 = sheet.getRange('Resources!A48:A61').getValues();
  var wordsRange2 = sheet.getRange('Resources!B48:B61').getValues();
  var wordsRange3 = sheet.getRange('Resources!C48:C66').getValues();
  var wordsRange4 = sheet.getRange('Resources!D48:D55').getValues();
  
  // Flatten each array of words (to avoid 2D arrays)
  var wordsArray1 = wordsRange1.map(function(row) { return row[0]; });
  var wordsArray2 = wordsRange2.map(function(row) { return row[0]; });
  var wordsArray3 = wordsRange3.map(function(row) { return row[0]; });
  var wordsArray4 = wordsRange4.map(function(row) { return row[0]; });
  
  // Call the function to search for words in the target text for each set of words
  var result1 = searchMultipleWords(wordsArray1, targetText);
  var result2 = searchMultipleWords(wordsArray2, targetText);
  var result3 = searchMultipleWords(wordsArray3, targetText);
  var result4 = searchMultipleWords(wordsArray4, targetText);
  
  // Output the results in cells C1, C2, C3, and C4
  /*sheet.getRange('AI Prompt!B8').setValue(result1); 
  sheet.getRange('AI Prompt!B9').setValue(result2);
  sheet.getRange('AI Prompt!B10').setValue(result3);
  sheet.getRange('AI Prompt!B11').setValue(result4);*/

  // Combine all results into one string
  var combinedResults = "";
  if (result1) combinedResults += result1 + "\n";
  if (result2) combinedResults += result2 + "\n";
  if (result3) combinedResults += result3 + "\n";
  if (result4) combinedResults += result4 + "\n";
  
  // Output the combined results into a single cell, e.g., AI Prompt!B8
  sheet.getRange('AI Prompt!AI4').setValue(combinedResults);
}

function runSearchMATSCI5() {
  //var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Get the target text from cell B1 (you can change this to any target cell)
  var targetText = sheet.getRange('AI Prompt!AO4').getValue();
  
  // Get words/phrases from 4 different ranges:
  var wordsRange1 = sheet.getRange('Resources!A48:A61').getValues();
  var wordsRange2 = sheet.getRange('Resources!B48:B61').getValues();
  var wordsRange3 = sheet.getRange('Resources!C48:C66').getValues();
  var wordsRange4 = sheet.getRange('Resources!D48:D55').getValues();
  
  // Flatten each array of words (to avoid 2D arrays)
  var wordsArray1 = wordsRange1.map(function(row) { return row[0]; });
  var wordsArray2 = wordsRange2.map(function(row) { return row[0]; });
  var wordsArray3 = wordsRange3.map(function(row) { return row[0]; });
  var wordsArray4 = wordsRange4.map(function(row) { return row[0]; });
  
  // Call the function to search for words in the target text for each set of words
  var result1 = searchMultipleWords(wordsArray1, targetText);
  var result2 = searchMultipleWords(wordsArray2, targetText);
  var result3 = searchMultipleWords(wordsArray3, targetText);
  var result4 = searchMultipleWords(wordsArray4, targetText);
  
  // Output the results in cells C1, C2, C3, and C4
  /*sheet.getRange('AI Prompt!B8').setValue(result1); 
  sheet.getRange('AI Prompt!B9').setValue(result2);
  sheet.getRange('AI Prompt!B10').setValue(result3);
  sheet.getRange('AI Prompt!B11').setValue(result4);*/

  // Combine all results into one string
  var combinedResults = "";
  if (result1) combinedResults += result1 + "\n";
  if (result2) combinedResults += result2 + "\n";
  if (result3) combinedResults += result3 + "\n";
  if (result4) combinedResults += result4 + "\n";
  
  // Output the combined results into a single cell, e.g., AI Prompt!B8
  sheet.getRange('AI Prompt!AS4').setValue(combinedResults);
}

function runSearchMATSCI6() {
  //var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Get the target text from cell B1 (you can change this to any target cell)
  var targetText = sheet.getRange('AI Prompt!AY4').getValue();
  
  // Get words/phrases from 4 different ranges:
  var wordsRange1 = sheet.getRange('Resources!A48:A61').getValues();
  var wordsRange2 = sheet.getRange('Resources!B48:B61').getValues();
  var wordsRange3 = sheet.getRange('Resources!C48:C66').getValues();
  var wordsRange4 = sheet.getRange('Resources!D48:D55').getValues();
  
  // Flatten each array of words (to avoid 2D arrays)
  var wordsArray1 = wordsRange1.map(function(row) { return row[0]; });
  var wordsArray2 = wordsRange2.map(function(row) { return row[0]; });
  var wordsArray3 = wordsRange3.map(function(row) { return row[0]; });
  var wordsArray4 = wordsRange4.map(function(row) { return row[0]; });
  
  // Call the function to search for words in the target text for each set of words
  var result1 = searchMultipleWords(wordsArray1, targetText);
  var result2 = searchMultipleWords(wordsArray2, targetText);
  var result3 = searchMultipleWords(wordsArray3, targetText);
  var result4 = searchMultipleWords(wordsArray4, targetText);
  
  // Output the results in cells C1, C2, C3, and C4
  /*sheet.getRange('AI Prompt!B8').setValue(result1); 
  sheet.getRange('AI Prompt!B9').setValue(result2);
  sheet.getRange('AI Prompt!B10').setValue(result3);
  sheet.getRange('AI Prompt!B11').setValue(result4);*/

  // Combine all results into one string
  var combinedResults = "";
  if (result1) combinedResults += result1 + "\n";
  if (result2) combinedResults += result2 + "\n";
  if (result3) combinedResults += result3 + "\n";
  if (result4) combinedResults += result4 + "\n";
  
  // Output the combined results into a single cell, e.g., AI Prompt!B8
  sheet.getRange('AI Prompt!BC4').setValue(combinedResults);
}

function searchCheck() {
  // Get the spreadsheet and specific sheet
  //var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName("Landing Page"); // Replace "SheetName" with the actual sheet name
  
  if (!sheet) {
    throw new Error("Sheet not found.");
  }

  // Get the value of the target cell
  var targetCell = sheet.getRange("C4").getValue(); // Replace "A1" with the desired cell reference

  // Check if the cell contains the word "CLEAR"
  if (targetCell.toString().trim().toUpperCase() === "MATSCI") {
    runSearchMATSCI(); // Replace with your function name
    runSearchMATSCI2(); // Replace with your function name
    runSearchMATSCI3(); // Replace with your function name
    runSearchMATSCI4(); // Replace with your function name
    runSearchMATSCI5(); // Replace with your function name
    runSearchMATSCI6(); // Replace with your function name
  }

  if (targetCell.toString().trim().toUpperCase() === "HUMANITIES") {
    runSearch(); // Replace with your function name
    runSearch2(); // Replace with your function name
    runSearch3(); // Replace with your function name
    runSearch4(); // Replace with your function name
    runSearch5(); // Replace with your function name
    runSearch6(); // Replace with your function name
  }
}