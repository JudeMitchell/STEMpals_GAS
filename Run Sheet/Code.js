function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('STEMpals')
      .addItem('Create School Folders', 'FileManagement.createSchoolFolders')
      .addToUi();
}

function createSchoolFolders_runSheet(){
  const ui = SpreadsheetApp.getUi()
  FileManagement.createSchoolFolders(ui);
}