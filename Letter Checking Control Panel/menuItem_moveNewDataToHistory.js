function moveNewDataToHistory() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  HelperFunctions.moveCheckerDataToHistory('safety', ss);
}
