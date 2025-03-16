function myFunction() {
  var ui = SpreadsheetApp.getUi();
  
  // Show custom "Running" message
  ui.showToast("⏳ Loading...");

  try {
    // 🚀 Your main script logic
    Utilities.sleep(3000); // Simulate a process
  } catch (error) {
    ui.showToast("❌ Error: " + error.message);
    return;
  }

  // Show custom "Script Finished" message
  ui.showToast("✅ Script Finished Successfully!");
}
