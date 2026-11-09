function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('STEMpals')
      .addItem('Generate student analysis forms', 'generateStudentAnalysisForms')
      .addSeparator()
      .addItem('Transform data from staging', 'transformFromStaging')
      .addToUi();
      
}