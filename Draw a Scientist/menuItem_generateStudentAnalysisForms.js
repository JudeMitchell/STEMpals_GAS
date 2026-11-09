function generateStudentAnalysisForms() {
  // the new files have to go into the folders that are shared with the teachers
  
  var testing = false;
  
  const analysisSpreadsheet = SpreadsheetApp.openById('1JQUN8i9dopAW8RrTfc0NrRzq_sJ-39wR-4-twibFnYk');
  const controlPanelTab = analysisSpreadsheet.getSheetByName('Control Panel');
  const schoolListTab = analysisSpreadsheet.getSheetByName('School List');
  const stempalsLogoID = '1sBwWXV2t2cQJhNOlNUGV0OynkFoK7K8x';
  const stempalsLogoBlob = DriveApp.getFileById(stempalsLogoID).getBlob();

  var studentFormID = controlPanelTab.getRange('B2').getValue();
  var templateResponseSSID = controlPanelTab.getRange('B4').getValue();
  var allSchoolsFolderID = controlPanelTab.getRange('B6').getValue();

  var allSchoolsFolder = DriveApp.getFolderById(allSchoolsFolderID);
  var schoolFolders = allSchoolsFolder.getFolders();

  var classes = {};

  while ( schoolFolders.hasNext() ){
    var sf = schoolFolders.next();
    if ( /Microsoft/.test(sf.getName()) ){
      continue;
    } else if ( /STEMpals[0-9]{2}/.test(sf.getName()) ) {
      var schoolName = sf.getName().replace(/STEMpals[0-9]{2}_/, '');
      
      var schoolClassRoomCount = schoolListTab.getRange('A2:A').createTextFinder(schoolName).findAll().length;
      if ( schoolClassRoomCount == 1){
        classes[schoolName] = [sf];
      } else if ( schoolClassRoomCount > 1) {
        classes[schoolName] = {};

        var classFolders = sf.getFolders();
        while ( classFolders.hasNext() ){
          var cf = classFolders.next();
          if ( ! /[R|r]esources/.test(cf.getName()) & ! /Term/.test(cf.getName()) & ! /Pairing/.test(cf.getName()) & ! /Hannah/.test(cf.getName()) ){
            classes[schoolName][cf.getName()] = cf;
          }
        }
      }
    }
  }
  // classes['cheese'] = {bacon: DriveApp.getFolderById('1iF6Iw__E-_YkAsRgBVz4JlfLemgrjgP2'), kabana: DriveApp.getFolderById('1iF6Iw__E-_YkAsRgBVz4JlfLemgrjgP2')}
  if ( testing ){
    Logger.log(classes);
  }

  var schoolDataClmnInds = HelperFunctions.findAllColumns('index', schoolListTab, 'School Name')
  var schoolsData = schoolListTab.getRange(2, 1, schoolListTab.getMaxRows() - 1, schoolListTab.getMaxColumns()).getValues();
  var emailDetails = {};

  for ( i = 0; i < schoolsData.length; i++ ){
    analysisSpreadsheet.toast('Making form ' + Number(i + 1) +  ' of ' + schoolsData.length);
    var iSchoolEmail = schoolsData[i][schoolDataClmnInds['Teacher Email']];
    var iSchoolName = schoolsData[i][schoolDataClmnInds['School Name']];
    var iOfficialClassName = schoolsData[i][schoolDataClmnInds['Class Official Name']];
    if ( testing ){
      Logger.log('i loop: ' + iSchoolName);
    }

    if ( schoolsData[i][schoolDataClmnInds['Code Name']] == '' ){
      var iClassName = iOfficialClassName;
    } else {
      var iClassName = iOfficialClassName + ' - ' + schoolsData[i][schoolDataClmnInds['Code Name']];
    }

    if ( ! Object.keys(emailDetails).includes(iSchoolEmail)) {
      emailDetails[iSchoolEmail] = {teacherName: schoolsData[i][schoolDataClmnInds['Teacher Name']], classes: {}};
    }

    if ( Array.isArray(classes[iSchoolName]) ){
      // single class school
      var destinationFolder = classes[iSchoolName][0];
    } else {
      var destinationFolder = classes[iSchoolName][iClassName];
    }

    if ( ! testing ){
      if ( schoolsData[i][schoolDataClmnInds['Form ID & Response Link']] != '' & schoolsData[i][schoolDataClmnInds['Form responses link']] != '' & schoolsData[i][schoolDataClmnInds['Response Spreadsheet']] != ''){
        var schoolFormID = schoolsData[i]['Form ID & Response Link'];
        var schoolForm = FormApp.openById(schoolFormID);
        var schoolResponseSS = DriveApp.getFileById(FormApp.openById(schoolFormID).getDestinationId());
      } else {
        var schoolFormID = DriveApp.getFileById(studentFormID).makeCopy(iSchoolName + ' ' + iOfficialClassName + ' Draw a Scientist Form', destinationFolder).getId();
        var schoolForm = FormApp.openById(schoolFormID).setPublished(true).setAcceptingResponses(true);
        schoolForm.addPublishedReader(iSchoolEmail);
        schoolForm.addEditor(iSchoolEmail);
        var schoolResponseSS = DriveApp.getFileById(templateResponseSSID).makeCopy(iSchoolName + ' ' + iOfficialClassName + ' Draw A Scientist Responses', destinationFolder);
        schoolForm.setDestination(FormApp.DestinationType.SPREADSHEET, schoolResponseSS.getId());
        schoolResponseSS.addEditor(iSchoolEmail);
      }
    
      var schoolFormLink = SpreadsheetApp.newRichTextValue().setText(schoolFormID).setLinkUrl(schoolForm.getPublishedUrl()).build();
      var schoolResultsLink = SpreadsheetApp.newRichTextValue().setText('Edit link').setLinkUrl(schoolForm.getEditUrl() + '#responses').build();
      var schoolSpreadsheetLink = SpreadsheetApp.newRichTextValue().setText('Spreadsheet responses').setLinkUrl(schoolResponseSS.getUrl()).build();
      
      if ( schoolsData[i][schoolDataClmnInds['Form ID & Response Link']] == '' ){
        schoolListTab.getRange('F' + Number(i + 2)).setRichTextValue(schoolFormLink);
      }
      if ( schoolsData[i][schoolDataClmnInds['Form responses link']] == '' ){
        schoolListTab.getRange('G' + Number(i + 2)).setRichTextValue(schoolResultsLink);
      }
      if ( schoolsData[i][schoolDataClmnInds['Response Spreadsheet']] == '' ){
        schoolListTab.getRange('H' + Number(i + 2)).setRichTextValue(schoolSpreadsheetLink);
      }
    }

    if ( ! testing ){
      emailDetails[iSchoolEmail]['classes'][iOfficialClassName] = {'responderLink': schoolForm.getPublishedUrl(),
                                                                  'formResponsesLink': schoolForm.getEditUrl() + '#responses',
                                                                  'spreadsheetLink': schoolResponseSS.getUrl(),
                                                                  'folderLink': destinationFolder.getUrl()};
    } else {
      emailDetails[iSchoolEmail]['classes'][iOfficialClassName] = { 'responderLink': 'responderLink', 
                                                                    'formResponsesLink': 'formResponsesLink',
                                                                    'spreadsheetLink': 'spreadsheetLink',
                                                                    'folderLink': 'folderLink'};
    }
  }

  var emailSubject =  'STEMpals: Draw a STEM professional analysis';
  var emailContentStart =  "Your <b>Draw-a-STEM professional</b> worksheets are now ready and available in your class folder. You’ll find everything you need there for your students to complete their final drawings for the year. Once your students have finished, you can complete the <b>online analysis form.</b><p>" +
                          "<b>Why do the analysis?</b>" +
                          "<ul>" +
                          "<li>It gives you a quick, visual way to see how your students’ views of STEM professionals have evolved since Term 1.</li>" + 
                          "<li>It supports student reflection and discussion around diversity and real-world STEM careers.</li>" + 
                          "<li>It provides data that you can use for a <b>maths or data analysis lesson</b> (optional!),  Access to your class’s results so students can explore and graph their findings.</li>" + 
                          "</ul>" +
                          "<p>" +
                          "<b>What to do:</b>";

  var emailOutro =  "This repeat activity is such an important way for us all to measure and celebrate how students’ perceptions of STEM professionals have grown this year.<p>" +
                    "Until then stay safe, sane, and sanitised, <br>" +
                    "Jen<p>" +
                    "<b>Jennifer Payne</b><br>" +
                    "Scientist & Founder<p>" +
                    '<a href="mailto:jen@STEMpals.org.au">jen@STEMpals.org.au</a><br>' +
                    "+61 437 353 827<br>" +
                    '<a href="https://www.linkedin.com/company/curiosityfactory">LinkedIn</></a><br>' +
                    "<img src=\'cid:stempalsLogo\'><p>";

  for ( email of Object.keys(emailDetails) ){
    var emailIntro = 'Hi ' + emailDetails[email]['teacherName'] + ',<p>';
    
    var emailInstructions = '<ol>';
    emailInstructions = emailInstructions + '<li>Have students complete their new “Draw-a-STEM professional” worksheets found in your class folder.<ul>';
    for ( className of Object.keys(emailDetails[email]['classes'])){
      emailInstructions = emailInstructions + '<li><a href="' + emailDetails[email]['classes'][className]['folderLink'] + '">Folder for class: ' + className + '</a><br></li>';
    }
    emailInstructions = emailInstructions + '</ul></li>';

    emailInstructions = emailInstructions + '<li>Scan and send the drawings back to me.</li>';
    
    emailInstructions = emailInstructions + '<li>Have your students complete the <b>short analysis form</b> here:<ul>';
    for ( className of Object.keys(emailDetails[email]['classes'])){
      emailInstructions = emailInstructions + '<li><a href="' + emailDetails[email]['classes'][className]['responderLink'] + '">Form for class: ' + className + '</a><br></li>';
    }
    emailInstructions = emailInstructions + '</ul></li>';

    emailInstructions = emailInstructions + '<li>See the visualisation of your classroom results here: <ul>';
    for ( className of Object.keys(emailDetails[email]['classes'])){
      emailInstructions = emailInstructions + '<li><a href="' + emailDetails[email]['classes'][className]['formResponsesLink'] + '">Visualisation for class: ' + className + '</a><br></li>';
    }
    emailInstructions = emailInstructions + '</ul></li>';

    emailInstructions = emailInstructions + '<li>(Optional) Use the <b>class results</b> for a data analysis activity.<ul>';
    for ( className of Object.keys(emailDetails[email]['classes'])){
      emailInstructions = emailInstructions + '<li><a href="' + emailDetails[email]['classes'][className]['spreadsheetLink'] + '">Data for class: ' + className + '</a><br></li>';
    }
    emailInstructions = emailInstructions + '</ul></li>';

    emailInstructions = emailInstructions + '</ol>';

    if ( ! testing ){
      MailApp.sendEmail({ 'to': email, 
                          'replyTo': 'jen@stempals.org.au',
                          'subject': emailSubject,
                          'htmlBody': emailIntro + emailContentStart + emailInstructions + emailOutro,
                          'inlineImages': {
                                            stempalsLogo: stempalsLogoBlob
                                          }
                          })
    }
  }
  Browser.msgBox('Forms generated & emails sent!!');
}


