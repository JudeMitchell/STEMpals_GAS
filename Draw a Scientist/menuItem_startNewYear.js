function startNewYear() {
  const analysisSSID = '1JQUN8i9dopAW8RrTfc0NrRzq_sJ-39wR-4-twibFnYk'
  const analysisSpreadsheet = SpreadsheetApp.openById(analysisSSID);
  const controlPanelTab = analysisSpreadsheet.getSheetByName('Control Panel');
  const schoolListTab = analysisSpreadsheet.getSheetByName('School List');
  

  
  //   // ARCHIVE A COPY FOR HISTORICAL PURPOSES
  // var lastYearAllSchoolsFolderID = controlPanelTab.getRange('B6').getValue();
  // var lastYearFolder = DriveApp.getFolderById(lastYearAllSchoolsFolderID);
  // var lastYear = lastYearFolder.getName().match(/[0-9]{4}/)[0];
  // DriveApp.getFileById(analysisSSID).makeCopy(lastYear + ' Draw a Scientist Analysis ARCHIVE', lastYearFolder);

    // CONTROL PANEL
  // ask for link to this year's school folder
  

  // SCHOOL LIST
  schoolListTab.getRange(2, 1, schoolListTab.getMaxRows() - 1, schoolListTab.getMaxColumns()).clearContent();

}
