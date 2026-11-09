function combineCheckerData(type, ss) {
  if ( type === 'safety' ){
    var dataRowsInputRange = 'B2:J';
    var dataLandingFirstColumn = 'C';
    var dataLandingLastColumn = 'K';
    var reviewFolderID = ss.getSheetByName('Control Panel').getRange('B4').getValue();
    var dataYear = ss.getSheetByName('Control Panel').getRange('B6').getValue();
    var dataLetterNumber = ss.getSheetByName('Control Panel').getRange('B7').getValue();
    var richTextInputColumn = 'B';
    var richTextLandingColumn = 'C';
    var dataValidationColumnShift = 3;
    var redFlagOutcomeColumn = 'M';
  } else { 
    //Draw a scientist
    // var templateSpreadsheetID = '1b28WskzKLsGWTC8NZq4PZ8Fvll3DIynnvtyEDHv3MM0';
    // var folderNameColumn = 'B';
    // var headerRowRange = 'A1:N1';
    // var inputRowRange = 'A2:N2';
    // var reviewFolderID = '1ByzB_KGRKH96xdD146BGlNl52X7_234y'; // jude's review playground
    // var reviewFolderID = '1ByzB_KGRKH96xdD146BGlNl52X7_234y';
  }

  var reviewFolder = DriveApp.getFolderById(reviewFolderID);
  var subFolders = reviewFolder.getFolders();
  var groupFolders = {};

  while (subFolders.hasNext()) {
    const folder = subFolders.next();
    var iFolderName = folder.getName();
    // Logger.log(iFolderName);
    
    if ( type === "safety"){
      groupFolders[iFolderName] = folder.getId();
    } else {
      if ( /group/.test(iFolderName) ){
        groupFolders[iFolderName] = folder.getId();
      }
    }
  }

  var newDataTab = ss.getSheetByName('New Data');

  if ( newDataTab.getFilter() != null ){  
    newDataTab.getFilter().remove();
  }

  for ( const groupfolderName of Object.keys(groupFolders) ){
    var groupFolder = DriveApp.getFolderById(groupFolders[groupfolderName]);

    var iGroupFolderFiles = groupFolder.getFiles();
    while ( iGroupFolderFiles.hasNext() ) {
      const file = iGroupFolderFiles.next();
      if ( file.getMimeType() === MimeType.GOOGLE_SHEETS ) {
        var iSS = SpreadsheetApp.openById(file.getId());
        var iSSTabs = iSS.getSheets();
        var iDataSheet = iSSTabs[0];

        var iData = iDataSheet.getRange(dataRowsInputRange + iDataSheet.getMaxRows()).getValues();
        var iRichTextData = iDataSheet.getRange(richTextInputColumn + '2:' + richTextInputColumn + iDataSheet.getMaxRows()).getRichTextValues();

        var newDataNextRow = newDataTab.getMaxRows() + 1;
        newDataTab.insertRowsAfter(newDataTab.getMaxRows(), iDataSheet.getMaxRows() - 1);

        newDataTab.getRange(dataLandingFirstColumn + newDataNextRow + ':' + dataLandingLastColumn + newDataTab.getMaxRows()).setValues(iData);
        for ( i = 0; i < iRichTextData.length; i++ ){
          newDataTab.getRange('A' + Number(i + newDataNextRow)).setValue(dataYear);
          newDataTab.getRange('B' + Number(i + newDataNextRow)).setValue(dataLetterNumber);
          newDataTab.getRange(richTextLandingColumn + Number(i + newDataNextRow)).setRichTextValue(iRichTextData[i][0]);
        }        
      }
    }
  }

  // copy data validation out of the template
  var templateSpreadsheetID = ss.getSheetByName('Control Panel').getRange('B2').getValue();
  var templateSheet = SpreadsheetApp.openById(templateSpreadsheetID).getSheetByName('data input');
  var templateDataValidations = templateSheet.getRange(dataRowsInputRange + '2').getDataValidations()[0];

  for ( j = 0; j < templateDataValidations.length; j++ ){
    if ( templateDataValidations[j] != null ){
      newDataTab.getRange(2, j+dataValidationColumnShift, newDataTab.getMaxRows() - 1, 1).setDataValidation(templateDataValidations[j]);
    }
  }

  if ( type == 'safety' ){
    var outcomeDV = SpreadsheetApp.newDataValidation().requireValueInList(['Redacted', 'OK in context', 'Serious']).build();
    newDataTab.getRange(redFlagOutcomeColumn + '2:' + redFlagOutcomeColumn + newDataTab.getMaxRows()).setDataValidation(outcomeDV);
  }

  newDataTab.getRange('A1:1').setFontWeight('bold');

  newDataTab.getRange(1, 1, newDataTab.getMaxRows(), newDataTab.getMaxColumns()).createFilter();
  Browser.msgBox('All done!!');
}
