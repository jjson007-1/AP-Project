/**
 * This function searches for a word from one cell in another cell (case-insensitive).
 * 
 * @param {string} wordToSearch The word to search for.
 * @param {string} targetText The text where we will search for the word.
 * @return {string} Returns a message whether the word was found or not.
 */
function searchWordInCell(wordToSearch, targetText) {
  // Convert both word and text to lower case to make the search case-insensitive.
  var lowerWord = wordToSearch.toLowerCase();
  var lowerTargetText = targetText.toLowerCase();
  
  // Check if the word is found in the target text using the `indexOf` method.
  if (lowerTargetText.indexOf(lowerWord) !== -1) {
    // If found, return a message saying the word was found.
    return "" +lowerWord;
  } else {
    // If not found, return a message saying the word was not found.
    return "";
  }
}

/**
 * Example: To use this script, you can call the function from a custom menu or trigger.
 * This function adds a custom menu to Google Sheets for easy execution.
 */
/*function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Word Search')
    .addItem('Search for Word', 'runSearch')  // Add a custom menu item to run the search
    .addToUi();
}*/

/**
 * This function runs the search and displays the result in the spreadsheet.
 */
function runSearch() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Get the word to search for from cell A1
  var wordToSearch = sheet.getRange('F6').getValue();
  
  // Get the target cell where the word will be searched, for example cell B1
  var targetText = sheet.getRange('A1').getValue();
  
  // Call the search function and show the result in cell C1
  var result = searchWordInCell(wordToSearch, targetText);
  
  // Output the result to cell C1
  sheet.getRange('C1').setValue(result);
}


