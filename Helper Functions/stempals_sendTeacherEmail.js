function sendTeacherEmails(emailType, emailDetails, emailVariables) {
  // emailDetails is a dict
  const stempalsLogoID = '1sBwWXV2t2cQJhNOlNUGV0OynkFoK7K8x';
  const stempalsLogoBlob = DriveApp.getFileById(stempalsLogoID).getBlob();

  var emailTypeConfig = {pairingSurvey: { emailSubject: 'STEMpals: Student pairing survey open',
                                          emailContentStart: "Your <b>Student Pairing Surveys</b> are now ready. Please have your students fill this out <b>before the {pairingSurveyDueDate}</b>.",
                                          emailInstructionsStart: '',
                                          emailOutro: 'This survey is used to help match your students with STEM professionals with similar interests.',
                                          signOff: "Until then stay safe, sane, and sanitised, "},
                         pairingSurveyFollowUp: { emailSubject: 'STEMpals: Student pairing survey incomplete',
                                                  emailContentStart: "Your <b>Student Pairing Surveys</b> are not yet complete. Please have your students fill this out <b>before the {pairingSurveyDueDate}</b>.",
                                                  emailInstructionsStart: '',
                                                  emailOutro: 'This survey is used to help match your students with STEM professionals with similar interests.',
                                                  signOff: "Until then stay safe, sane, and sanitised, "},
                         welcomePack: { emailSubject: 'Welcome to STEMpals ' + emailVariables['year'],
                                        emailContentStart: "Welcome to STEMpals " + emailVariables['year'] + "! 🎉<p>",
                                        emailInstructionsStart: '',
                                        emailOutro: "We're so excited to see where your students' curiosity takes them!",
                                        signOff: 'Stay curious,'}

                         };

  var emailSignature =  "<p>" + emailTypeConfig[emailType]['signOff'] + "<br>" +
                        "Jen<p>" +
                        "<b>Jennifer Payne</b><br>" +
                        "Scientist & Founder<p>" +
                        '<a href="mailto:jen@STEMpals.org.au">jen@STEMpals.org.au</a><br>' +
                        "+61 437 353 827<br>" +
                        '<a href="https://www.linkedin.com/company/curiosityfactory">LinkedIn</></a><p>' +
                        "<img src=\'cid:stempalsLogo\'><p>";

  for ( e of Object.keys(emailDetails) ){
    if ( Object.keys(emailDetails[e]['classes']).length == 0 ){
      continue;
    }
    Logger.log(emailDetails[e])
    var eEmailIntro = 'Hi ' + emailDetails[e]['teacherName'] + ',<p>';
  
    var eEmailInstructions = '';

    if ( emailType == 'pairingSurvey' ){
      eEmailInstructions = eEmailInstructions + "<br><ul>"
      for ( className of Object.keys(emailDetails[e]['classes'])){
        eEmailInstructions = eEmailInstructions + '<li><a href="' + emailDetails[e]['classes'][className]['responderLink'] + '">Form for class: ' + className + '</a><br></li>';
      }
      eEmailInstructions = eEmailInstructions + "</ul>"
    } else if ( emailType == 'pairingSurveyFollowUp' ){
      eEmailInstructions = eEmailInstructions + "<br><ul>"
      for ( className of Object.keys(emailDetails[e]['classes']) ){
        eEmailInstructions = eEmailInstructions + '<li><a href="' + emailDetails[e]['classes'][className]['formLink'] + '">Form for class: ' + className + '</a><br></li><ul>';
        for ( studentName of emailDetails[e]['classes'][className]['missingStudents'] ){
          eEmailInstructions = eEmailInstructions + '<li>' + studentName + '</li>';
        }
        eEmailInstructions = eEmailInstructions + '</ul>';
      }
      eEmailInstructions = eEmailInstructions + '</ul>';
    } else if ( emailType == 'welcomePack' ){
      eEmailInstructions = eEmailInstructions + "We're thrilled to have you and your students joining us!<p>";
      eEmailInstructions = eEmailInstructions + "Your <b>STEMpals folder</b> is ready with everything you need to get started.<p>";
      eEmailInstructions = eEmailInstructions + '📁 <a href="' + emailDetails[e]['folderLink'] + '">Access your STEMpals folder here</a><p><p>';
      eEmailInstructions = eEmailInstructions + '<b><p style="font-size: 18px;">What\'s inside?</b></p>';
      '<br><br><a href="' + emailDetails[e]['folderLink'] + '">Folder for your school: ' + emailDetails[e]['schoolName'] + '</a><p><p>';
      var welcomePackDetails = getWelcomePackContents();
      for ( wpd of welcomePackDetails ){
        eEmailInstructions = eEmailInstructions + '<li><b>' + wpd['filename'] + ': </b>' + wpd['fileInformation'] + ' </li>'
      }
      eEmailInstructions = eEmailInstructions + '</ul><p>';

      eEmailInstructions = eEmailInstructions + '<b><p style="font-size: 18px;">Your first two steps</b></p><ol>';
      var firstStepDetails = getFirstStepsContents(emailVariables);
      for ( fsd of firstStepDetails ){
        eEmailInstructions = eEmailInstructions + '<li><b>' + fsd['task'] + ': </b>' + fsd['instruction'] + ' </li>'
      }
      eEmailInstructions = eEmailInstructions + '</ol><p>';

      eEmailInstructions = eEmailInstructions + "Once your list is complete, I'll send through a short survey about your students' hobbies and interests to help us find their STEM professional matches! For students that don’t have guardian permission we will provide a generic letter from a STEM professional so they can still work along with their peers.<p>";

      eEmailInstructions = eEmailInstructions + '<b><p style="font-size: 18px;">Throughout the year</b></p>';
      eEmailInstructions = eEmailInstructions + "Your folder will also be where you upload scanned student letters and receive letters from STEM professionals, along with student pairings and updates. I'll guide you through each stage with reminders along the way.<p>";

    }
    
    eEmailInstructions = eEmailInstructions + "";

    var allEmailContent = eEmailIntro + emailTypeConfig[emailType]['emailContentStart'] + eEmailInstructions + emailTypeConfig[emailType]['emailOutro'] + emailSignature;
    
    Logger.log(allEmailContent)
    if ( Object.keys(emailVariables).length > 0 ){
      for ( ev of Object.keys(emailVariables) ){
        
        allEmailContent = allEmailContent.replace('{' + ev + '}', emailVariables[ev]);
      }
    }

    // MailApp.sendEmail({ 'to': 'judemitchell36@gmail.com', 
    MailApp.sendEmail({ 'to': e, 
                        'replyTo': 'jen@stempals.org.au',
                        'subject': emailTypeConfig[emailType]['emailSubject'],
                        'htmlBody': allEmailContent,
                        'inlineImages': {
                                          stempalsLogo: stempalsLogoBlob
                                        }
                        });
  }
}


function getWelcomePackContents(){
  var welcomePackContents = [
                              {'filename': 'Welcome Pack',
                               'fileInformation': 'Program overview, lesson plans and classroom ideas.'},
                              {'filename': 'Key Dates',
                               'fileInformation': 'Important deadlines for the year.'},
                              {'filename': 'Guardian Permission Slip',
                               'fileInformation': 'An editable template for your school.'},
                              {'filename': 'Student List',
                               'fileInformation': "A spreadsheet for your students' first names participating."},
                              {'filename': 'Term Printables',
                               'fileInformation': 'Worksheets and activities for each term.'},
                            ];
  return welcomePackContents;
}

function getFirstStepsContents(emailVariables){
  var firstStepsContents = [
                              {'task': 'Obtain guardian permission',
                               'instruction': 'using the editable template.'},
                              {'task': "Add your students' first names",
                               'instruction': 'to the Student List spreadsheet by <b>' + emailVariables['studentListDueDate'] + '</b>. These are the students participating in STEMpals with guardian permission.'},
                            ];
  return firstStepsContents;
}