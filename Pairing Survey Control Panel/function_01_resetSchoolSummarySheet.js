function resetSchoolSummarySheet() {
  const pairingSurveySS = SpreadsheetApp.openById('1Gbwcwc0JyyrWd49I0-cS2lwMB-ync_p9w9XB0lKbu48');
  const controlPanelTab = pairingSurveySS.getSheetByName('Control Panel');
  const summaryTab = pairingSurveySS.getSheetByName('School Summary');

  const templateFormID = controlPanelTab.getRange('B2').getValue();
  const yearFolderID = controlPanelTab.getRange('B4').getValue();
  var yearFolder = DriveApp.getFolderById(yearFolderID);
  const year = controlPanelTab.getRange('B6').getValue();

  const yearRunSheetSSFile = yearFolder.getFilesByName(year + ' STEMpals Run Sheet').next();
  const yearRunSheetSS = SpreadsheetApp.open(yearRunSheetSSFile);
  const teachersTab = yearRunSheetSS.getSheetByName('Teachers');

  const teachersTabPaidStudentsClmnIndex = STEMpalsHelperFunctions.findColumnIndex(teachersTab, 'number students paid for');

  // school folders
  var schoolFolders = STEMpalsHelperFunctions.getSchoolFolders(yearFolderID, 'main');
  STEMpalsHelperFunctions.resetSheetSize(summaryTab, Object.keys(schoolFolders).length + 1, 5);

  var summaryRow = 2;
  for ( sf of Object.keys(schoolFolders) ){
    STEMpalsHelperFunctions.setRichTextCell(schoolFolders[sf].getUrl(), summaryTab.getRange('A' + summaryRow), 'folder', text=schoolFolders[sf].getName());  
  }
  const schoolStudentsPaid = STEMpalsHelperFunctions.getDictFromRange(teachersTab.getRange(2, 1, teachersTab.getMaxRows(), teachersTab.getMaxColumns()), 0, teachersTabPaidStudentsClmnIndex);
  
  STEMpalsHelperFunctions.putDictInTableColumn(schoolStudentsPaid, summaryTab, 'A', 'B', 2);

  summaryTab.getRange('D1').setValue('Make Pairing Survey').setFontWeight('bold');
  summaryTab.getRange('E1').setValue('Pairing Survey ID').setFontWeight('bold');
  summaryTab.getRange('F1').setValue('Pairing Survey Responses').setFontWeight('bold');
  summaryTab.getRange('G1').setValue('Response Count').setFontWeight('bold');
}


