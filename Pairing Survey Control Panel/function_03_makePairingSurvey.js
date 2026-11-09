function makePairingSurveys() {
  const pairingSurveySS = SpreadsheetApp.openById('1Gbwcwc0JyyrWd49I0-cS2lwMB-ync_p9w9XB0lKbu48');
  const controlPanelTab = pairingSurveySS.getSheetByName('Control Panel');
  const summaryTab = pairingSurveySS.getSheetByName('School Summary');

  const templateFormID = controlPanelTab.getRange('B2').getValue();
  const yearFolderID = controlPanelTab.getRange('B4').getValue();
  var yearFolder = DriveApp.getFolderById(yearFolderID);
  const year = controlPanelTab.getRange('B6').getValue();
  const surveyDueDate = controlPanelTab.getRange('B8').getValue();

  // make a folder for the forms to live in
  var yearSchoolsFolder = yearFolder.getFoldersByName('Schools').next();
  var foldersInYSF = yearSchoolsFolder.getFoldersByName('Pairing Surveys');
  if ( foldersInYSF.hasNext() ){
    var pairingSurveyFolder = foldersInYSF.next();
  } else {
    var pairingSurveyFolder = DriveApp.createFolder('Pairing Surveys');
    pairingSurveyFolder.moveTo(yearSchoolsFolder);

    var msPairingSurveyFolder = DriveApp.createFolder('Microsoft Pairing Survey Data');
    msPairingSurveyFolder.moveTo(yearSchoolsFolder);    
  }

  const yearRunSheetSSFile = yearFolder.getFilesByName(year + ' STEMpals Run Sheet').next();
  const yearRunSheetSS = SpreadsheetApp.open(yearRunSheetSSFile);
  const teachersTab = yearRunSheetSS.getSheetByName('Teachers');

  const summaryTabMakeSurveyClmnLtr = STEMpalsHelperFunctions.findColumnLetter(summaryTab, 'Make Pairing Survey');
  // const summaryTabSurveyIDClmnLtr = STEMpalsHelperFunctions.findColumnLetter(summaryTab, 'Students Registered');
  const summaryTabSurveyIDClmnLtr = STEMpalsHelperFunctions.findColumnLetter(summaryTab, 'Pairing Survey ID');

  // school config
  var schoolFolders = STEMpalsHelperFunctions.getSchoolFolders(yearFolderID, 'main');
  var schoolEmails = STEMpalsHelperFunctions.getDictFromRange(teachersTab.getRange('A2:C'), 0, 2);
  var schoolTeachers = STEMpalsHelperFunctions.getDictFromRange(teachersTab.getRange('A2:G'), 0, 6);
  var schoolsWithNonUniqueNames = [];

  for ( summaryRow = 2; summaryRow <= summaryTab.getMaxRows(); summaryRow++ ){
    var emailDetails = {};
    var srMakeSurvey = summaryTab.getRange(summaryTabMakeSurveyClmnLtr + summaryRow).getValue();
    if ( srMakeSurvey ){
      var iSurveyIDs = [];
      var iSchoolFolder = schoolFolders[summaryTab.getRange('A' + summaryRow).getValue()];
      var iSchoolName = iSchoolFolder.getName().replace('STEMpals', '').replace(year.toString().slice(2,4), '').replace('_', '').replace(' - Microsoft', '');
      var iSchoolEmail = schoolEmails[iSchoolFolder];
      var iSchoolTeacher = schoolTeachers[iSchoolFolder];

      if ( ! Object.keys(emailDetails).includes(iSchoolEmail)) {
        emailDetails[iSchoolEmail] = {teacherName: iSchoolTeacher, classes: {}};
      }

      var sfFiles = iSchoolFolder.getFiles();
      while ( sfFiles.hasNext() ){
        var sff = sfFiles.next()
        
        if ( sff.getName() == "[Temp] Student List" ){
          sff.setTrashed(true);
        } else if ( /Student.*[l|L]ist/.test(sff.getName()) | /Class list/.test(sff.getName()) ){
          if ( sff.getMimeType() != MimeType.GOOGLE_SHEETS ){
            var sfStudentListBlob = sff.getBlob();
            const config = {
                            name: "[Temp] Student List",
                            mimeType: MimeType.GOOGLE_SHEETS,
                            parents: [iSchoolFolder.getId()]
                          };
            var sfStudentListConverted = Drive.Files.create(config, sfStudentListBlob);
            var sfStudentListSS = SpreadsheetApp.openById(sfStudentListConverted.id);
          } else {
            var sfStudentListSS = SpreadsheetApp.open(sff);
          } 
          var sfStudentListTab = sfStudentListSS.getSheetByName('Sheet1');
          if ( sfStudentListTab.getRange('A3').getValue() == '' ){
            summaryTab.getRange(summaryTabStudentsRegisteredClmnLtr + summaryRow).setValue(0);
          } else {
            var sfSendSurvey = false;
            // check for first name duplicates
            var sfNumStudents = STEMpalsHelperFunctions.getListFromRange(sfStudentListTab.getRange('A3:A'), 'column').length;

            sfStudentListTab.getRange(2, 4, 1, 1).setValue("First Name Count").setFontWeight('bold');
            sfStudentListTab.getRange(3, 4, sfNumStudents, 1).setFormula("=countif(A$3:A,A3)");
            sfStudentListTab.getRange(2, 5, 1, 1).setValue("Student Code").setFontWeight('bold');
            sfStudentListTab.getRange(3, 5, sfNumStudents, 1).setFormula('=if(D3=1,A3,A3&" "&B3)');
            sfStudentListTab.getRange(2, 6, 1, 1).setValue("Student Code in School Count").setFontWeight('bold');
            sfStudentListTab.getRange(3, 6, sfNumStudents, 1).setFormula("=countifs(E$3:E,E3)");

            var studentCodeInSchoolCounts = STEMpalsHelperFunctions.getDistinctListFromRange(sfStudentListTab.getRange(3, 6, sfNumStudents, 1), 'column');
            if ( studentCodeInSchoolCounts.filter(nbr => nbr > 1).length > 0 ){
              sfStudentListTab.getRange(2, 7, 1, 1).setValue("Student/Class Code").setFontWeight('bold');
              sfStudentListTab.getRange(3, 7, sfNumStudents, 1).setFormula('=if(F3=1,E3,E3&" "&C3)');
              sfStudentListTab.getRange(2, 8, 1, 1).setValue("Student/Class Code Count").setFontWeight('bold');
              sfStudentListTab.getRange(3, 8, sfNumStudents, 1).setFormula("=countif(G$3:G,G3)");
              var studentClassCodeInSchoolCounts = STEMpalsHelperFunctions.getDistinctListFromRange(sfStudentListTab.getRange(3, 8, sfNumStudents, 1), 'column');
              if ( studentClassCodeInSchoolCounts.filter(nbr => nbr > 1).length > 0 ){
                schoolsWithNonUniqueNames.push(iSchoolName);
                summaryTab.getRange('E' + summaryRow).setBackground('red').setValue('Non Unique Names');
              } else {
                sfSendSurvey = true;
                var classStudentLists = STEMpalsHelperFunctions.getDictFromRange(sfStudentListTab.getRange(3, 3, sfNumStudents, 5), 0, 4, true);
              }
            } else {
              sfSendSurvey = true;
              var classStudentLists = STEMpalsHelperFunctions.getDictFromRange(sfStudentListTab.getRange(3, 3, sfNumStudents, 3), 0, 2, true);
            }

            if ( sfSendSurvey ){
              // var classStudentLists = STEMpalsHelperFunctions.getDictFromRange(sfStudentListTab.getRange(3, 3, sfNumStudents, 3), 0, 2, true);

              for ( className of Object.keys(classStudentLists) ){
                var schoolFormID = DriveApp.getFileById(templateFormID).makeCopy(iSchoolName + ' ' + className + ' STEMpals Pairing', pairingSurveyFolder).getId();
                var schoolForm = FormApp.openById(schoolFormID);
                schoolForm.setTitle(iSchoolName + ' ' + className + ' STEMpals Pairing');
                var items = schoolForm.getItems();
                var studentNameItem = items[0].asListItem();
                studentNameItem.setChoiceValues(classStudentLists[className].sort());

                schoolForm.addPublishedReader(iSchoolEmail);
                Drive.Permissions.create({'type': 'anyone', 'view': 'published', 'role': 'reader'}, schoolForm.getId());

                emailDetails[iSchoolEmail]['classes'][className] = {'responderLink': schoolForm.getPublishedUrl()};
                schoolForm.setPublished(true).setAcceptingResponses(true);
                iSurveyIDs.push(schoolFormID);
              }
            }
          }
        }
      }
      if ( iSurveyIDs.length > 0 ){
        summaryTab.getRange(summaryTabSurveyIDClmnLtr + summaryRow).setValue(iSurveyIDs.join('\n')).setBackground(null);
        summaryTab.getRange(summaryTabMakeSurveyClmnLtr + summaryRow).setValue(false);
      }
    }
    STEMpalsHelperFunctions.sendTeacherEmails('pairingSurvey', emailDetails, {pairingSurveyDueDate: surveyDueDate});
  }

  var checkboxDV = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  summaryTab.getRange('D2:D').setDataValidation(checkboxDV);
  
  if( schoolsWithNonUniqueNames.length > 0 ){
    Browser.msgBox('The following schools had issues with non unique names:\\n' + schoolsWithNonUniqueNames.join('\\n - '));
  }
}
