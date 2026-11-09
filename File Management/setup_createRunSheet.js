function createRunSheet() {
  const ui = SpreadsheetApp.getUi();
  
  const fileManagerSS = SpreadsheetApp.openById('1zO0JTbchfqLHyYTG6PYleFsGzSrOByOphvwqThV3Fwg');
  const controlPanelSheet = fileManagerSS.getSheetByName('Control Panel');
  const userGuideSheet = fileManagerSS.getSheetByName('User Guide');
  const runSheetIDsSheet = fileManagerSS.getSheetByName('Run Sheet IDs');
  const folderIDsSheet = fileManagerSS.getSheetByName('Folder IDs');
  const runSheetTemplate = getTemplateFile('Run Sheet', 'file');

  // copy link formulas
  var userGuideFormulas = userGuideSheet.getRange('B1:C2').getFormulas();

  // get year details
  var year = controlPanelSheet.getRange('B1').getValue();
  var yearFolderID = controlPanelSheet.getRange('B3').getValue();
  var yearFolder = DriveApp.getFolderById(yearFolderID);

    ////// CREATE RUNSHEET
  var yearRunSheetID = HelperFunctions.runSheetTemplate.makeCopy(year + " STEMpals Run Sheet", yearFolder).getId();

    ////// UPDATE RUN SHEET IDS TAB
  runSheetIDsSheet.insertRowBefore(2);
  runSheetIDsSheet.getRange('B2').setValue(yearRunSheetID);
  HelperFunctions.setRichTextCell(yearRunSheetID, runSheetIDsSheet.getRange('A2'), 'spreadsheet', text=year);
  HelperFunctions.setRichTextCell(yearRunSheetID, runSheetIDsSheet.getRange('C2'), 'spreadsheet', text=year + ' Run sheet');
  
    ////// UPDATE YEAR FOLDERS TAB
  folderIDsSheet.insertRowBefore(2);
  folderIDsSheet.getRange('B2').setValue(yearFolder.getId());
  HelperFunctions.setRichTextCell(yearFolder.getId(), folderIDsSheet.getRange('A2'), 'folder', text=year);
  HelperFunctions.setRichTextCell(yearFolder.getId(), folderIDsSheet.getRange('C2'), 'folder', text=year + ' Folder');

    ////// MAKE THE EMPTY FOLDERS
  HelperFunctions.updateFolder('YYYY STEMpals Program', yearFolder, year);
  
    ////// REPLACE LINK FORMULAS after row insertion messes them up
  userGuideSheet.getRange('B1:C2').setFormulas(userGuideFormulas);

  ui.alert(year + ' run sheet and folder created!', 'Please add the Welcome Pack pdf to the folder and fill in the Teachers tab of the run sheet.', ui.ButtonSet.OK)
}
