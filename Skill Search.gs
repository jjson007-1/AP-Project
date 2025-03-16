/**
 * This function searches for multiple words from multiple cells in a target cell (case-insensitive).
 * It checks each word or phrase in the source cells and searches it in the target cell.
 * 
 * @param {Array} wordsArray An array of words/phrases to search for.
 * @param {string} targetText The text where we will search for the words/phrases.
 * @return {string} Returns a message listing all words that were found or not found.
 */
function searchMultipleSkill(wordsArray, targetText) {
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
        result += '' + word + '';  // Add the found word to the result
      } else {
        result += '';  // Add the not found word to the result
      }
    }
  });
  
  return result;  // Return the result string
}

/**
 * This function is triggered from the custom menu to run the search for multiple words.
 */

var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
function runSearchSkill() {
  //var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Get the target text from cell (you can change this to any target cell)
  var targetText = sheet.getRange('AI Prompt!A7').getValue();
  
  // Get words/phrases from 4 different ranges:
  var wordsRange1 = sheet.getRange('Resources!AA2:AA11').getValues();
  var wordsRange2 = sheet.getRange('Resources!AD2:AD7').getValues();
  var wordsRange3 = sheet.getRange('Resources!AG2:AG9').getValues();
  var wordsRange4 = sheet.getRange('Resources!AJ2:AJ9').getValues();
  
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
  /*var combinedResults = "";
  if (result1) combinedResults += result1 + ""; 
  if (result2) combinedResults += result2 + ""; 
  if (result3) combinedResults += result3 + "";
  if (result4) combinedResults += result4 + ""; */

  // Combine all results into one string, filtering out empty results
  var results = [result1, result2, result3, result4].filter(function(result) {
    return result && result.trim() !== ""; // Ensure no empty or whitespace-only results
  });

  var combinedResults = results.join(", "); // Join results with a comma and space
  
  // Output the combined results into a single cell, e.g., AI Prompt!B8
  sheet.getRange('AI Prompt!E7').setValue(combinedResults);
}


function runSearchSkill2() {
  //var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Get the target text from cell (you can change this to any target cell)
  var targetText = sheet.getRange('AI Prompt!K7').getValue();
  
  // Get words/phrases from 4 different ranges:
  var wordsRange1 = sheet.getRange('Resources!AA2:AA11').getValues();
  var wordsRange2 = sheet.getRange('Resources!AD2:AD7').getValues();
  var wordsRange3 = sheet.getRange('Resources!AG2:AG9').getValues();
  var wordsRange4 = sheet.getRange('Resources!AJ2:AJ9').getValues();
  
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
  /*var combinedResults = "";
  if (result1) combinedResults += result1 + ""; 
  if (result2) combinedResults += result2 + ""; 
  if (result3) combinedResults += result3 + "";
  if (result4) combinedResults += result4 + ""; */

  // Combine all results into one string, filtering out empty results
  var results = [result1, result2, result3, result4].filter(function(result) {
    return result && result.trim() !== ""; // Ensure no empty or whitespace-only results
  });

  var combinedResults = results.join(", "); // Join results with a comma and space
  
  // Output the combined results into a single cell, e.g., AI Prompt!B8
  sheet.getRange('AI Prompt!O7').setValue(combinedResults);
}


function runSearchSkill3() {
  //var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Get the target text from cell (you can change this to any target cell)
  var targetText = sheet.getRange('AI Prompt!U7').getValue();
  
  // Get words/phrases from 4 different ranges:
  var wordsRange1 = sheet.getRange('Resources!AA2:AA11').getValues();
  var wordsRange2 = sheet.getRange('Resources!AD2:AD7').getValues();
  var wordsRange3 = sheet.getRange('Resources!AG2:AG9').getValues();
  var wordsRange4 = sheet.getRange('Resources!AJ2:AJ9').getValues();
  
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
  /*var combinedResults = "";
  if (result1) combinedResults += result1 + ""; 
  if (result2) combinedResults += result2 + ""; 
  if (result3) combinedResults += result3 + "";
  if (result4) combinedResults += result4 + ""; */

  // Combine all results into one string, filtering out empty results
  var results = [result1, result2, result3, result4].filter(function(result) {
    return result && result.trim() !== ""; // Ensure no empty or whitespace-only results
  });

  var combinedResults = results.join(", "); // Join results with a comma and space
  
  // Output the combined results into a single cell, e.g., AI Prompt!B8
  sheet.getRange('AI Prompt!Y7').setValue(combinedResults);
}


function runSearchSkill4() {
  //var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Get the target text from cell (you can change this to any target cell)
  var targetText = sheet.getRange('AI Prompt!AE7').getValue();
  
  // Get words/phrases from 4 different ranges:
  var wordsRange1 = sheet.getRange('Resources!AA2:AA11').getValues();
  var wordsRange2 = sheet.getRange('Resources!AD2:AD7').getValues();
  var wordsRange3 = sheet.getRange('Resources!AG2:AG9').getValues();
  var wordsRange4 = sheet.getRange('Resources!AJ2:AJ9').getValues();
  
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
  /*var combinedResults = "";
  if (result1) combinedResults += result1 + ""; 
  if (result2) combinedResults += result2 + ""; 
  if (result3) combinedResults += result3 + "";
  if (result4) combinedResults += result4 + ""; */

  // Combine all results into one string, filtering out empty results
  var results = [result1, result2, result3, result4].filter(function(result) {
    return result && result.trim() !== ""; // Ensure no empty or whitespace-only results
  });

  var combinedResults = results.join(", "); // Join results with a comma and space
  
  // Output the combined results into a single cell, e.g., AI Prompt!B8
  sheet.getRange('AI Prompt!AI7').setValue(combinedResults);
}


function runSearchSkill5() {
  //var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Get the target text from cell (you can change this to any target cell)
  var targetText = sheet.getRange('AI Prompt!AO7').getValue();
  
  // Get words/phrases from 4 different ranges:
  var wordsRange1 = sheet.getRange('Resources!AA2:AA11').getValues();
  var wordsRange2 = sheet.getRange('Resources!AD2:AD7').getValues();
  var wordsRange3 = sheet.getRange('Resources!AG2:AG9').getValues();
  var wordsRange4 = sheet.getRange('Resources!AJ2:AJ9').getValues();
  
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
  /*var combinedResults = "";
  if (result1) combinedResults += result1 + ""; 
  if (result2) combinedResults += result2 + ""; 
  if (result3) combinedResults += result3 + "";
  if (result4) combinedResults += result4 + ""; */

  // Combine all results into one string, filtering out empty results
  var results = [result1, result2, result3, result4].filter(function(result) {
    return result && result.trim() !== ""; // Ensure no empty or whitespace-only results
  });

  var combinedResults = results.join(", "); // Join results with a comma and space
  
  // Output the combined results into a single cell, e.g., AI Prompt!B8
  sheet.getRange('AI Prompt!AS7').setValue(combinedResults);
}


function runSearchSkill6() {
  //var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Get the target text from cell (you can change this to any target cell)
  var targetText = sheet.getRange('AI Prompt!AY7').getValue();
  
  // Get words/phrases from 4 different ranges:
  var wordsRange1 = sheet.getRange('Resources!AA2:AA12').getValues();
  var wordsRange2 = sheet.getRange('Resources!AD2:AD7').getValues();
  var wordsRange3 = sheet.getRange('Resources!AG2:AG9').getValues();
  var wordsRange4 = sheet.getRange('Resources!AJ2:AJ9').getValues();
  
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
 /* var combinedResults = "";
  if (result1) combinedResults += result1 + ""; 
  if (result2) combinedResults += result2 + ""; 
  if (result3) combinedResults += result3 + "";
  if (result4) combinedResults += result4 + ""; */

  // Combine all results into one string, filtering out empty results
  var results = [result1, result2, result3, result4].filter(function(result) {
    return result && result.trim() !== ""; // Ensure no empty or whitespace-only results
  });

  var combinedResults = results.join(", "); // Join results with a comma and space
  
  // Output the combined results into a single cell, e.g., AI Prompt!B8
  sheet.getRange('AI Prompt!BC7').setValue(combinedResults); 
}


function runSearchPhil() {
  //var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Get the target text from cell (you can change this to any target cell)
  var targetText = sheet.getRange('Lesson Plan AI Prompt!D2').getValue();
  
  // Get words/phrases from 4 different ranges:
  var wordsRange1 = sheet.getRange('Resources!AM17:AM32').getValues();
  var wordsRange2 = sheet.getRange('Resources!AD50:AD57').getValues();
  var wordsRange3 = sheet.getRange('Resources!AG50:AG57').getValues();
  var wordsRange4 = sheet.getRange('Resources!AJ50:AJ57').getValues();
  
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
 /* var combinedResults = "";
  if (result1) combinedResults += result1 + ""; 
  if (result2) combinedResults += result2 + ""; 
  if (result3) combinedResults += result3 + "";
  if (result4) combinedResults += result4 + ""; */

  // Combine all results into one string, filtering out empty results
  var results = [result1, result2, result3, result4].filter(function(result) {
    return result && result.trim() !== ""; // Ensure no empty or whitespace-only results
  });

  var combinedResults = results.join(", "); // Join results with a comma and space
  
  // Output the combined results into a single cell, e.g., AI Prompt!B8
  sheet.getRange('Landing Page!E2').setValue(combinedResults); 
}
/**
 * This function adds a custom menu to run the search.
 */

