function checkForRegisteredStudents() {
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
  const summaryTabPaidStudentsClmnLtr = STEMpalsHelperFunctions.findColumnLetter(summaryTab, 'Students Paid For');
  const summaryTabStudentsRegisteredClmnLtr = STEMpalsHelperFunctions.findColumnLetter(summaryTab, 'Students Registered');

  // school folders
  var schoolFolders = STEMpalsHelperFunctions.getSchoolFolders(yearFolderID, 'main');
  var schoolFoldersInTable = STEMpalsHelperFunctions.getListFromRange(summaryTab.getRange('A2:A'), 'column');

  // var summaryRow = 2;
  // var schoolsToAdd = {};
  for ( sf of Object.keys(schoolFolders) ){
    Logger.log(sf);
    if ( ! schoolFoldersInTable.includes(sf) ){
      summaryTab.insertRowAfter(summaryTab.getMaxRows());
      summaryRow = summaryTab.getMaxRows();
      STEMpalsHelperFunctions.setRichTextCell(schoolFolders[sf].getUrl(), summaryTab.getRange('A' + summaryRow), 'folder', text=schoolFolders[sf].getName());

      const schoolStudentsPaid = STEMpalsHelperFunctions.getDictFromRange(teachersTab.getRange(2, 1, teachersTab.getMaxRows(), teachersTab.getMaxColumns()), 0, teachersTabPaidStudentsClmnIndex);
      summaryTab.getRange(summaryTabPaidStudentsClmnLtr + summaryRow).setValue(schoolStudentsPaid[sf]);
    } else {
      var summaryRow = schoolFoldersInTable.indexOf(sf) + 2;
    }
    
    var sfFiles = schoolFolders[sf].getFiles();
    if ( ! sfFiles.hasNext() ){
      summaryTab.getRange(summaryTabStudentsRegisteredClmnLtr + summaryRow).setValue('no files in folder');
    }
    while ( sfFiles.hasNext() ){
      var sff = sfFiles.next()
      
      if ( sff.getName() == "[Temp] Student List" ){
        sff.setTrashed(true);
      } else if ( ! /pdf/.test(sff.getName()) & (/Student.*[l|L]ist/.test(sff.getName()) | /Class list/.test(sff.getName())) ){
        var sfStudentListSS = openAnyFileAsSS(sff, schoolFolders[sf], "Student List");
        // if ( sff.getMimeType() != MimeType.GOOGLE_SHEETS ){
        //   var sfStudentListBlob = sff.getBlob();
        //   const config = {
        //                   name: "[Temp] Student List",
        //                   mimeType: MimeType.GOOGLE_SHEETS,
        //                   parents: [schoolFolders[sf].getId()]
        //                 };
        //   var sfStudentListConverted = Drive.Files.create(config, sfStudentListBlob);
        //   var sfStudentListSS = SpreadsheetApp.openById(sfStudentListConverted.id);
        // } else {
        //   var sfStudentListSS = SpreadsheetApp.open(sff);
        // } 
        var sfStudentListTab = sfStudentListSS.getSheetByName('Sheet1');
        if ( sfStudentListTab.getRange('A3').getValue() == '' ){
          summaryTab.getRange(summaryTabStudentsRegisteredClmnLtr + summaryRow).setValue(0);
        } else {
          var sfNumStudents = STEMpalsHelperFunctions.getListFromRange(sfStudentListTab.getRange('A3:A'), 'column').length;
          summaryTab.getRange(summaryTabStudentsRegisteredClmnLtr + summaryRow).setValue(sfNumStudents);
        }
      }
    }
  }

  // var registeredRange = summaryTab.getRange('C2:C');
  // var allRegisteredCF = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=C2=B2').setBackground('#34a853').setRanges([registeredRange]).build(); //green
  // var marginOfErrorRegisteredCF = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=and(C2>=B2*0.9,C2<=B2*1.1)').setBackground('#93c47d').setRanges([registeredRange]).build(); //lighter green
  // var noneRegisteredCF = SpreadsheetApp.newConditionalFormatRule().whenNumberEqualTo(0).setBackground('#ff6d01').setRanges([registeredRange]).build(); //orange
  // var notEnoughRegisteredCF = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied("=C2<B2").setBackground('#fbbc04').setRanges([registeredRange]).build(); //yellow
  // var noFilesCF = SpreadsheetApp.newConditionalFormatRule().whenTextContains('no files').setBackground('#ff6d01').setRanges([registeredRange]).build(); //orange
  // var tooManyRegisteredCF = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied("=C2>B2").setBackground('#cc0100').setRanges([registeredRange]).build() //red
  // summaryTab.setConditionalFormatRules([allRegisteredCF, marginOfErrorRegisteredCF, noneRegisteredCF, notEnoughRegisteredCF, noFilesCF, tooManyRegisteredCF]);
  updateFormatting(summaryTab);

  var checkboxDV = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  summaryTab.getRange('D2:D').setDataValidation(checkboxDV);
  summaryTab.getRange('D2:D').setFormula('=if(and(C2>=B2*0.9,C2<=B2*1.1,C2<>"no files in folder",E2=""),TRUE,FALSE)');
  var makeSurveyDefaultValues = summaryTab.getRange('D2:D').getValues();
  summaryTab.getRange('D2:D').setValues(makeSurveyDefaultValues);

}


