function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('STEMpals')
      .addItem('Generate checker spreadsheets', 'generateCheckerSpreadsheets')
      .addSeparator()
      .addItem('Combine checker data', 'combineCheckerSpreadsheets')
      .addSeparator()
      .addItem('Move New Data to History', 'moveNewDataToHistory')
      .addToUi();
      
}