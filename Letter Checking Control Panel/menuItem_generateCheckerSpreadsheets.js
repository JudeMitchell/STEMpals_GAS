function generateCheckerSpreadsheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  HelperFunctions.generateCheckerSpreadsheets('safety', ss);
}
