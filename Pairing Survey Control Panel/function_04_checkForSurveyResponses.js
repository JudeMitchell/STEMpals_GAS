function checkForSurveyResponses() {
  const pairingSurveySS = SpreadsheetApp.openById('1Gbwcwc0JyyrWd49I0-cS2lwMB-ync_p9w9XB0lKbu48');
  const controlPanelTab = pairingSurveySS.getSheetByName('Control Panel');
  const responsesTab = pairingSurveySS.getSheetByName('Survey Responses');
  const expectedResponsesTab = pairingSurveySS.getSheetByName('Expected Responses');
  const summaryTab = pairingSurveySS.getSheetByName('School Summary');

  const yearFolderID = controlPanelTab.getRange('B4').getValue();
  const yearFolder = DriveApp.getFolderById(yearFolderID);
  const year = controlPanelTab.getRange('B6').getValue();
  const surveyDueDate = controlPanelTab.getRange('B8').getValue();
  
  const yearRunSheetSSFile = yearFolder.getFilesByName(year + ' STEMpals Run Sheet').next();
  const yearRunSheetSS = SpreadsheetApp.open(yearRunSheetSSFile);
  const teachersTab = yearRunSheetSS.getSheetByName('Teachers');

  const schoolsFolder = yearFolder.getFoldersByName('Schools').next();
  const microsoftDataFolder = schoolsFolder.getFoldersByName('Microsoft Pairing Survey Data').next();

  const summaryTabSurveyIDClmnLtr = STEMpalsHelperFunctions.findColumnLetter(summaryTab, 'Pairing Survey ID');
  const summaryTabMissingResponsesClmnLtr = STEMpalsHelperFunctions.findColumnLetter(summaryTab, 'Missing Responses');
  const summaryTabResponseSendEmailClmnLtr = STEMpalsHelperFunctions.findColumnLetter(summaryTab, 'Send Follow Up Email');

  var formResponderLinks = {};
  var responseData = [];
  var microsoftData = [];
  var microsoftSchools = {};
  var tempFilesCreated = [];
  var schoolsToEmail = [];

  for ( sRow = 2; sRow <= summaryTab.getMaxRows(); sRow ++ ){
    var srIDString = summaryTab.getRange(summaryTabSurveyIDClmnLtr + sRow).getValue();

    if ( srIDString == 'Microsoft' ){
      var schoolName = summaryTab.getRange('A' + sRow).getValue().slice(11).replace(' - Microsoft', '');
      if ( summaryTab.getRange(summaryTabResponseSendEmailClmnLtr + sRow).getValue() ){
        schoolsToEmail.push(schoolName);
      }
      
      microsoftSchools[schoolName] = summaryTab.getRange('A' + sRow).getValue();

      var schoolNameRegex = new RegExp(schoolName);
      var mpsdFiles = microsoftDataFolder.getFiles();
      while ( mpsdFiles.hasNext() ){
        var mpsdFile = mpsdFiles.next();
        if ( /Temp/.test(mpsdFile.getName()) ){
          if ( ! mpsdFile.isTrashed() ){
            mpsdFile.setTrashed(true);
          }
        } else if ( schoolNameRegex.test(mpsdFile.getName()) ){
          // convert the csv for Survey Responses tab
          var schoolMPSDBlob = mpsdFile.getBlob();
          const config = {
                          name: "[Temp] " + mpsdFile.getName(),
                          mimeType: MimeType.GOOGLE_SHEETS,
                          parents: [microsoftDataFolder.getId()]
                        };
          var schoolMPSDConverted = Drive.Files.create(config, schoolMPSDBlob);
          tempFilesCreated.push(schoolMPSDConverted);
          var schoolMPSDSS = SpreadsheetApp.openById(schoolMPSDConverted.id);
          var schoolMPSDSheet = schoolMPSDSS.getSheets()[0];

          var smpsdDataCount = STEMpalsHelperFunctions.getListFromRange(schoolMPSDSheet.getRange('A2:A'), 'column').length;
          
          schoolMPSDSheet.createTextFinder('Mark on the scale how strongly you agree or disagree with each question..').matchEntireCell(false).replaceAllWith('');
          schoolMPSDSheet.createTextFinder('Who is better at science').matchEntireCell(true).replaceAllWith('Who is better at science?');
          schoolMPSDSheet.createTextFinder('I understand what scientists do in their job').matchEntireCell(true).replaceAllWith('I understand what scientists do in their jobs');
          schoolMPSDSheet.createTextFinder('I would be interested in learning about different jobs in science').matchEntireCell(true).replaceAllWith('I would be interested in learning about different types of jobs in science');
          
          schoolMPSDSheet.createTextFinder('Strongly agree (YES!)').matchEntireCell(true).replaceAllWith('Strongly agree');
          schoolMPSDSheet.createTextFinder('agree').matchEntireCell(true).replaceAllWith('Agree');
          schoolMPSDSheet.createTextFinder('neutral').matchEntireCell(true).replaceAllWith('Neutral');
          schoolMPSDSheet.createTextFinder('disagree').matchEntireCell(true).replaceAllWith('Disagree');
          schoolMPSDSheet.createTextFinder('Strongly disagree (NO!)').matchEntireCell(true).replaceAllWith('Strongly disagree');
          
          schoolMPSDSheet.createTextFinder('NO!').matchEntireCell(true).replaceAllWith('Strongly disagree');
          schoolMPSDSheet.createTextFinder('no').matchEntireCell(true).replaceAllWith('Disagree');
          schoolMPSDSheet.createTextFinder('Maybe').matchEntireCell(true).replaceAllWith('Neutral');
          schoolMPSDSheet.createTextFinder('yes').matchEntireCell(true).replaceAllWith('Agree');
          schoolMPSDSheet.createTextFinder('YES!').matchEntireCell(true).replaceAllWith('Strongly agree');
          
          schoolMPSDSheet.createTextFinder(1).matchEntireCell(true).replaceAllWith('Women are much better');
          schoolMPSDSheet.createTextFinder(2).matchEntireCell(true).replaceAllWith('Women are slightly better');
          schoolMPSDSheet.createTextFinder(3).matchEntireCell(true).replaceAllWith('No difference');
          schoolMPSDSheet.createTextFinder(4).matchEntireCell(true).replaceAllWith('Men are slightly better');
          schoolMPSDSheet.createTextFinder(5).matchEntireCell(true).replaceAllWith('Men are much better');
          
          var selectThingsYouEnjoyClmnLtr = schoolMPSDSheet.createTextFinder('Select things you enjoy').matchEntireCell(false).matchCase(false).findAll()[0].getA1Notation().replace('1', '');
          var petsClmnLtr = schoolMPSDSheet.createTextFinder('If you have any pets').matchEntireCell(false).matchCase(false).findAll()[0].getA1Notation().replace('1', '');
          
          for ( stuRow = 2; stuRow <= smpsdDataCount + 1; stuRow++ ){
            var thingsEnjoyedList = schoolMPSDSheet.getRange(selectThingsYouEnjoyClmnLtr + stuRow).getValue();
            if ( thingsEnjoyedList.slice(thingsEnjoyedList.length - 1) == ';' ){
              schoolMPSDSheet.getRange(selectThingsYouEnjoyClmnLtr + stuRow).setValue(thingsEnjoyedList.slice(0, thingsEnjoyedList.length - 1));
            }

            var petsCell = schoolMPSDSheet.getRange(petsClmnLtr + stuRow);
            if ( petsCell.getValue() == 'Disagree' ){
              petsCell.setValue('no');
            }
          }
          schoolMPSDSheet.createTextFinder(';').matchEntireCell(false).replaceAllWith(', ');

          var sheetData = {'schoolName': schoolName, 'columns': {}};
          for ( iClmn = 1; iClmn <= schoolMPSDSheet.getMaxRows(); iClmn++ ){
            var iClmnName = schoolMPSDSheet.getRange(1, iClmn, 1, 1).getValue();
            sheetData['columns'][iClmnName.toUpperCase()] = schoolMPSDSheet.getRange(2, iClmn, smpsdDataCount, 1).getValues();
          }
          microsoftData.push(sheetData);

          const smpsdName2ClmnLtr = STEMpalsHelperFunctions.findColumnLetter(schoolMPSDSheet, 'Name2');
          // srResponseCount += STEMpalsHelperFunctions.getDistinctListFromRange(schoolMPSDSheet.getRange(smpsdName2ClmnLtr + '2:' + smpsdName2ClmnLtr), 'column').length;
        } // END OF SINGLE MICROSOFT FILE (where school name is in the file)
      } // END OF LOOPING OVER MPSD FOLDER FILES
      
      const summTabMicrosoftURLClmnLtr = STEMpalsHelperFunctions.findColumnLetter(summaryTab, 'Microsoft Pairing Survey URL');

      // there is only one survey per school for Microsoft
      if ( ! Object.keys(formResponderLinks).includes(schoolName) ){
        formResponderLinks[schoolName] = {};
      } 
      formResponderLinks[schoolName]['Microsoft'] = summaryTab.getRange(summTabMicrosoftURLClmnLtr + sRow).getRichTextValue().getLinkUrl();
      
      // END OF MICROSOFT SCHOOL LOGIC
    } else if ( srIDString != '' ){
      var srIDs = srIDString.split('\n');
      for ( fID of srIDs ){
        var fForm = FormApp.openById(fID);
        var fFormNameComponents = fForm.getTitle().split(' ');
        var fResponses = fForm.getResponses();

        var items = fForm.getItems();
        var studentNameItem = items[0].asListItem();
        
        for ( fr of fResponses ){
          var frResponse = fFormNameComponents.slice(0,2);
          var frItemResponses = fr.getItemResponses();
          for ( frir of frItemResponses ){
            if ( Array.isArray(frir.getResponse()) ){
              frResponse.push(frir.getResponse().join(', '))
            } else {
              frResponse.push(frir.getResponse())
            }
          }
          responseData.push(frResponse)
        }

        var fFormSchool = fFormNameComponents[0];

        if ( summaryTab.getRange(summaryTabResponseSendEmailClmnLtr + sRow).getValue() ){
          schoolsToEmail.push(fFormSchool);  
        }
        
        var fFormClass = fFormNameComponents[1];
        if ( ! Object.keys(formResponderLinks).includes(fFormSchool) ){
          formResponderLinks[fFormSchool] = {};
        } 
        if ( ! Object.keys(formResponderLinks[fFormSchool]).includes(fFormClass) ){
          formResponderLinks[fFormSchool][fFormClass] = fForm.getPublishedUrl();
        }

      }
    } // END OF GOOGLE SHEETS LOGIC
  }

  summaryTab.getRange(summaryTabMissingResponsesClmnLtr + '2:' + summaryTabMissingResponsesClmnLtr).setFormula('=countifs(\'Expected Responses\'!A$2:A,substitute(right(A2,len(A2)-11)," - Microsoft",""),\'Expected Responses\'!D$2:D,"0")');

  if ( responseData.length > 0 ){
    STEMpalsHelperFunctions.resetSheetData(responsesTab, 2, 1, responseData);
    
    // find the gender column and replace row by row
    var genderQuestionClmnNbr = responsesTab.createTextFinder('Who is better at science').matchEntireCell(false).matchCase(false).findAll()[0].getColumn();
    
    var numbers = [1, 2, 3, 4, 5];
    var genderValues = ['Women are much better', 'Women are slightly better', 'No difference', 'Men are slightly better', 'Men are much better'];
    var agreeValues = ['Strongly disagree', 'Disagree', 'Neutral', 'Agree', 'Strongly agree'];
    for ( i = 0; i < 5; i++ ){
      var nbrFinder = responsesTab.createTextFinder(numbers[i]).matchEntireCell(true).findAll();
      for ( nf of nbrFinder ){
        if ( nf.getColumn() == genderQuestionClmnNbr ){
          nf.setValue(genderValues[i]);
        } else {
          nf.setValue(agreeValues[i]);
        }
      }
    }
  } else {
    if ( responsesTab.getMaxRows() > 1 ){
      responsesTab.deleteRows(2, responsesTab.getMaxRows() - 1);
    }
  }

  if ( microsoftData.length > 0 ){
    for ( md of microsoftData ){
      var mdRowCount = md['columns']['NAME'].length;
      var rTabRow = responsesTab.getMaxRows() + 1;
      responsesTab.insertRowsAfter(responsesTab.getMaxRows(), mdRowCount);
      responsesTab.getRange(rTabRow, 1, mdRowCount, 1).setValue(md['schoolName']);

      for ( rtClmn = 2; rtClmn <= responsesTab.getMaxColumns(); rtClmn++ ){
        var rtClmnName = responsesTab.getRange(1, rtClmn, 1, 1).getValue();
        if ( rtClmnName == 'Name' ){
          responsesTab.getRange(rTabRow, rtClmn, mdRowCount, 1).setValues(md['columns']['NAME2']);
        } else if ( Object.keys(md['columns']).includes(rtClmnName.toUpperCase()) ){
          responsesTab.getRange(rTabRow, rtClmn, mdRowCount, 1).setValues(md['columns'][rtClmnName.toUpperCase()]);
        }
      }
    }
  }

  if ( responsesTab.getMaxRows() > 2 ){
    responsesTab.getRange(2, 1, responsesTab.getMaxRows() - 1, responsesTab.getMaxColumns()).setBackground(null).setFontWeight('normal');
  }

  expectedResponsesTab.getRange('D2:D').setFormula("=countifs('Survey Responses'!A$2:A,A2,'Survey Responses'!B$2:B,B2,'Survey Responses'!C$2:C,C2)");
  updateFormatting(summaryTab);

  const ui = SpreadsheetApp.getUi();
  var sendTeacherEmail = ui.alert('Would you like to send an email to teachers where responses are missing?', ui.ButtonSet.YES_NO);

  if ( sendTeacherEmail === ui.Button.YES ){
  // if ( true ){
    var schoolEmails = STEMpalsHelperFunctions.getDictFromRange(teachersTab.getRange('A2:C'), 0, 2);
    var schoolTeachers = STEMpalsHelperFunctions.getDictFromRange(teachersTab.getRange('A2:G'), 0, 6);
    var schoolFolders = STEMpalsHelperFunctions.getSchoolFolders(yearFolderID, 'main', true);

    var emailDetails = {};
    for ( erRow = 2; erRow <= expectedResponsesTab.getMaxRows(); erRow++ ){
      var erResponseCount = expectedResponsesTab.getRange('D' + erRow).getValue();
      if ( erResponseCount == 0 ){
        var erSchool = expectedResponsesTab.getRange('A' + erRow).getValue();
        if ( ! schoolsToEmail.includes(erSchool) ){
          continue;
        }
        var erClass = expectedResponsesTab.getRange('B' + erRow).getValue();
        var erStudent = expectedResponsesTab.getRange('C' + erRow).getValue();
        if ( Object.keys(microsoftSchools).includes(erSchool) ){
          var erSchoolKey = microsoftSchools[erSchool]  
          var erFormLink = formResponderLinks[erSchool]['Microsoft'];
        } else {
          var erSchoolKey = schoolFolders[erSchool];
          var erFormLink = formResponderLinks[erSchool][erClass];
        }
        
        var erSchoolEmail = schoolEmails[erSchoolKey];
        var erSchoolTeacher = schoolTeachers[erSchoolKey];
        
        if ( ! Object.keys(emailDetails).includes(erSchoolEmail) ){
          emailDetails[erSchoolEmail] = {teacherName: erSchoolTeacher, classes: {}}
        }
        if ( ! Object.keys(emailDetails[erSchoolEmail]['classes']).includes(erClass) ){
          emailDetails[erSchoolEmail]['classes'][erClass] = {formLink: erFormLink, missingStudents: []};
        }
        emailDetails[erSchoolEmail]['classes'][erClass]['missingStudents'].push(erStudent);
      }
    }

    STEMpalsHelperFunctions.sendTeacherEmails('pairingSurveyFollowUp', emailDetails, {pairingSurveyDueDate: surveyDueDate});
  }

  if ( tempFilesCreated.length > 0 ){
    for ( tfc of tempFilesCreated ){
      DriveApp.getFileById(tfc['id']).setTrashed(true);
    }
  }
}
