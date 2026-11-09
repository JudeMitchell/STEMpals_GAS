function generateCheckerSpreadsheets(type, ss) {
  if ( type === 'safety' ){
    // var templateSpreadsheetID = '1cj05GzwecqAH5s5wlZnPK4RvO1eLi8hCroLaf4l5NGM';
    var folderNameColumn = 'C'; //called Classroom but it'll do
    var headerRowRange = 'A1:L1';
    var inputRowRange = 'A2:L2';
    var reviewFolderID = ss.getSheetByName('Control Panel').getRange('B4').getValue();
  } else { 
    //Draw a scientist
    var templateSpreadsheetID = '1b28WskzKLsGWTC8NZq4PZ8Fvll3DIynnvtyEDHv3MM0';
    var folderNameColumn = 'B';
    var headerRowRange = 'A1:N1';
    var inputRowRange = 'A2:N2';
    // var reviewFolderID = '1ByzB_KGRKH96xdD146BGlNl52X7_234y'; // jude's review playground
    // var reviewFolderID = '1ByzB_KGRKH96xdD146BGlNl52X7_234y';
  }

  // shouldn't have to change any of this :)
  var reviewFolder = DriveApp.getFolderById(reviewFolderID);
  var templateSpreadsheetID = ss.getSheetByName('Control Panel').getRange('B2').getValue();

  var subFolders = reviewFolder.getFolders();
  var groupFolders = {};

  while (subFolders.hasNext()) {
    const folder = subFolders.next();
    var iFolderName = folder.getName();
    
    if ( type === "safety"){
      groupFolders[iFolderName] = folder.getId();
    } else {
      if ( /group/.test(iFolderName) ){
        groupFolders[iFolderName] = folder.getId();
      }
    }
  }

  var boldStyle = SpreadsheetApp.newTextStyle().setBold(true).build();

  var foldersWithNoPDFs = [];

  for ( const groupfolderName of Object.keys(groupFolders) ){
    var groupFolder = DriveApp.getFolderById(groupFolders[groupfolderName]);
    var groupSpreadsheet = DriveApp.getFileById(templateSpreadsheetID).makeCopy(groupfolderName, groupFolder);

    var iGroupFolderFiles = groupFolder.getFiles();
    var groupPDFs = {};
    while ( iGroupFolderFiles.hasNext() ) {
      const file = iGroupFolderFiles.next();
      if ( /pdf/.test(file.getName()) ) {
        groupPDFs[file.getName()] = file.getUrl();
      }
    }

    if ( Object.keys(groupPDFs).length > 0 ){
      var groupDataSheet = SpreadsheetApp.openById(groupSpreadsheet.getId()).getSheetByName('data input').activate();
      Logger.log(Object.keys(groupPDFs));
      if ( Object.keys(groupPDFs).length > 1 ){
        groupDataSheet.insertRowsAfter(2, Object.keys(groupPDFs).length - 1);   
      }
      var inputRow = groupDataSheet.getRange(inputRowRange);

      for ( i = 0; i < Object.keys(groupPDFs).length; i++ ) {
        var iFile = Object.keys(groupPDFs)[i];
        inputRow.copyTo(groupDataSheet.getRange('A' + Number(2 + i)));
        if ( type === 'safety' ){
          // for safety extract student name from filename
          var iRichText = SpreadsheetApp.newRichTextValue().setText(iFile.match(/(.*)-([0-9]+-[0-9]+.pdf)/)[1]).setLinkUrl(groupPDFs[iFile]).build();
          groupDataSheet.getRange('B' + Number(2 + i)).setRichTextValue(iRichText);
        } else {
          // Draw A Scientist gets full filename
          groupDataSheet.getRange('A' + Number(2 + i)).setValue(groupPDFs[i]);
        }
        
        groupDataSheet.getRange(folderNameColumn + Number(2 + i)).setValue(groupfolderName);
      }
      groupDataSheet.getRange(headerRowRange).setTextStyle(boldStyle);
    } else {
      foldersWithNoPDFs.push(groupfolderName);
    }
  }

  if ( foldersWithNoPDFs.length > 0 ){
    const ui = SpreadsheetApp.getUi();
    ui.alert('Some folders had no PDFs', " - " + foldersWithNoPDFs.join("\n - "), ui.ButtonSet.OK);
  }

  Browser.msgBox('All done!!');
}