function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('STEMpals')
      .addItem('Check for Registered Students', 'checkForRegisteredStudents') 
      .addItem('Make Pairing Surveys', 'makePairingSurveys') 
      .addItem('Check for Survey Responses', 'checkForSurveyResponses') 
      .addSeparator()
      .addItem('Reset School Summary Sheet', 'resetSchoolSummarySheet') 
      .addToUi();
}
