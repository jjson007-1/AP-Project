
function mainFunction() {
  Logger.log("Main function started.");

  
  checkAndExecuteSub();


  // Call function from Word Search.gs
  searchCheck();

  insertAIFormula2();

  // Call function from Transformation.js
  runSearchSkill();
  runSearchSkill2();
  runSearchSkill3();
  runSearchSkill4();
  runSearchSkill5();
  runSearchSkill6();

  // Call function from Reporting.js
  matchAndPrintMultipleRangesIgnoreSpaces();
  matchAndPrintMultipleRangesIgnoreSpaces2();
  matchAndPrintMultipleRangesIgnoreSpaces3();
  matchAndPrintMultipleRangesIgnoreSpaces4();
  matchAndPrintMultipleRangesIgnoreSpaces5();
  matchAndPrintMultipleRangesIgnoreSpaces6();

  matchAndPrintFromTwoSheets();
  matchAndPrintFromTwoSheets2();
  matchAndPrintFromTwoSheets3();
  matchAndPrintFromTwoSheets4();
  matchAndPrintFromTwoSheets5();
  matchAndPrintFromTwoSheets6();

  //checkAndExecute();
  //insertAIFormulaExPlan();
  //insertAIFormulaPrePlan();
  Logger.log("Main function completed.");
}

function LessonPlan() {
  checkAndExecute();
  insertAIFormulaExPlan();
  insertAIFormulaPrePlan();
}

function OverrideLessonPlan() {
  insertLessonPlanFormula();
  insertAIFormulaExPlan();
  insertAIFormulaPrePlan();
}

function LessonPlanContinuation() {
  insertLessonPlanFormulaCont();
  insertAIFormulaExPlan();
  insertAIFormulaPrePlan();
}

function SearchPhil() {
  runSearchPhil();
}

function SavePlan() {
  savePlan();
  exportToNewGoogleDoc();
}

function addSavePlanMenu(ui) {
  ui.createMenu('Save')
    .addItem('Save Lesson Plan', 'SavePlan')  
    .addToUi();
}

function addGenerateMenu(ui) {
  ui.createMenu('Set Up Plan')
    .addItem('Set Up Plan', 'mainFunction')  
    .addToUi();
}

function addGeneratePlanMenu(ui) {
  ui.createMenu('Create Plan') 
    .addItem('Create Lesson Plan', 'LessonPlan')  
    .addToUi();
}

function addOverrideMenu(ui) {
  ui.createMenu('Override and Create Plan') 
    .addItem('Create Lesson Plan', 'OverrideLessonPlan')
    .addItem('Create Continuation Lesson Plan', 'LessonPlanContinuation')  
    .addToUi();
}

//Menu Item for Report Function (See reportForm.gs)

function addReportMenu(ui) {
  ui.createMenu('Report Issue')
  .addItem('Submit Report', 'openReportForm')
  .addToUi();
}

/**
 * This function adds a custom menu to run the search.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  addGenerateMenu(ui);
  addGeneratePlanMenu(ui);
  addSavePlanMenu(ui);
  addReportMenu(ui);
  addOverrideMenu(ui);
}