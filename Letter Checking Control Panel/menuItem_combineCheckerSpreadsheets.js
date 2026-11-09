function combineCheckerSpreadsheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  HelperFunctions.combineCheckerData('safety', ss);
}
