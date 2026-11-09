function createSchoolFolders(ui) {
  if ( ui == null ){
    var ui = SpreadsheetApp.getUi();
  }
  
  const runSheetSS = HelperFunctions.getRunSheet();
  const teachersSheet = runSheetSS.getSheetByName('Teachers');
  const yearFolder = HelperFunctions.getYearFolder();
  const year = yearFolder.getName().slice(0,4);
  const yearSchoolsFolder = HelperFunctions.getYearFolder(null, 'Schools');

  HelperFunctions.checkForInputFiles('STEMpalsYY_<SchoolName>', 'WelcomePack', year, ui);

  var studentListDue = HelperFunctions.getKeyDate('Student List Due', '<EEEE> <d> <mmmm>');

  var failedSchools = [];
  var emailDetails = {};
  var teachersSheetClmnLtrs = HelperFunctions.findAllColumns('letter', teachersSheet, 'School');

  var emailVariables = {'year': year, 'studentListDueDate': studentListDue};

  // sort the teachers tab to deal with blank rows
  teachersSheet.getRange(4, 1, teachersSheet.getMaxRows() - 3, teachersSheet.getMaxColumns()).sort(1);

  // conditional formatting to remind Jen to send teachers Microsoft stuff
  var teachersSheetCFRules = teachersSheet.getConditionalFormatRules();

  var nSchools = HelperFunctions.getListFromRange(teachersSheet.getRange('A4:A'), 'column').length;
  for ( iRow = 4; iRow <= teachersSheet.getMaxRows(); iRow++ ){
    try {
      var iSchool = teachersSheet.getRange(teachersSheetClmnLtrs['School'] + iRow).getValue();
      if ( iSchool == '' ){
        break;
      }

      var iFolder = teachersSheet.getRange(teachersSheetClmnLtrs['Folder'] + iRow).getValue();
      if ( iFolder == '' | iFolder == '---' ){
        runSheetSS.toast(iSchool, 'School ' + Number(iRow - 3) + '/' + nSchools);
      } else {
        Logger.log (' - Skipping school with existing folder: ' + iSchool + '...');
        continue;
      } 
      
      var iMicrosoft = teachersSheet.getRange(teachersSheetClmnLtrs['Microsoft'] + iRow).getValue(); // allowed to be blank

      var iEmail = HelperFunctions.getNonBlankValue(teachersSheet.getRange(teachersSheetClmnLtrs['Editor Email'] + iRow), 'email address');
      var iTeacher = HelperFunctions.getNonBlankValue(teachersSheet.getRange(teachersSheetClmnLtrs['Teacher'] + iRow), 'teacher name');
      
      var iSchoolFolderName = 'STEMpals' + year.slice(2,4) + '_' + iSchool.replace(' ', '');
      if ( iMicrosoft ){
        iSchoolFolderName = iSchoolFolderName + ' - Microsoft';
      }
      var iSchoolFolder = yearSchoolsFolder.createFolder(iSchoolFolderName);

      HelperFunctions.updateFolder('STEMpalsYY_<SchoolName>', iSchoolFolder, year, 'RunSheet');

      HelperFunctions.setRichTextCell(iSchoolFolder.getId(), teachersSheet.getRange(teachersSheetClmnLtrs['Folder'] + iRow), 'folder', iSchoolFolderName);
      
      if ( ! iMicrosoft ){
        iSchoolFolder.addEditor(iEmail);

        emailDetails[iEmail] = {'classes': {'key': 'value'}, 
                                'teacherName': iTeacher, 
                                'folderLink': iSchoolFolder.getUrl(),
                                'schoolName': iSchool};
        HelperFunctions.sendTeacherEmails('welcomePack', emailDetails, emailVariables);
        teachersSheet.getRange(teachersSheetClmnLtrs['School'] + iRow).setBackground(null);
        teachersSheet.getRange(teachersSheetClmnLtrs['Welcome pack email'] + iRow).setValue(true);
      } else {
        var microsoftSentCF = SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(false).setBackground('DarkOrange').setRanges([teachersSheet.getRange(teachersSheetClmnLtrs['Welcome pack email'] + iRow)]).build();
        teachersSheetCFRules.push(microsoftSentCF);
      }

    } catch ( schoolError ){
      Logger.log('ERROR: ' + iSchool);
      Logger.log(schoolError.stack)
      teachersSheet.getRange(teachersSheetClmnLtrs['School'] + iRow).setBackground('crimson');
      failedSchools.push(iSchool);
    }
  }

  teachersSheet.setConditionalFormatRules(teachersSheetCFRules);

  if ( failedSchools.length > 0 ){
    ui.alert('Some schools failed!', '- ' + failedSchools.join('\n- '), ui.ButtonSet.OK);
  }
}
