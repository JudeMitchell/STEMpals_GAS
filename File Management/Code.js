function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('STEMpals')
    .addItem('Create Run Sheet', 'createRunSheet')
    .addToUi();
}



