function sendTeacherEmails(emailType, emailDetails, emailVariables) {
  // emailDetails is a dict
  const stempalsLogoID = '1sBwWXV2t2cQJhNOlNUGV0OynkFoK7K8x';
  const stempalsLogoBlob = DriveApp.getFileById(stempalsLogoID).getBlob();

  var emailTypeConfig = {pairingSurvey: { emailSubject: 'STEMpals: Student pairing survey open',
                                          emailContentStart: "Your <b>Student Pairing Surveys</b> are now ready. Please have your students fill this out <b>before the {pairingSurveyDueDate}</b>.",
                                          emailInstructionsStart: '',
                                          emailOutro: 'This survey is used to help match your students with STEM professionals with similar interests.'},
                         pairingSurveyFollowUp: { emailSubject: 'STEMpals: Student pairing survey incomplete',
                                                  emailContentStart: "Your <b>Student Pairing Surveys</b> are not yet complete. Please have your students fill this out <b>before the {pairingSurveyDueDate}</b>.",
                                                  emailInstructionsStart: '',
                                                  emailOutro: 'This survey is used to help match your students with STEM professionals with similar interests.'},
                         welcomePack: { emailSubject: 'STEMpals: Welcome to ' + emailVariables['year'] + '!',
                                        emailContentStart: "Thank you for signing up for a year of curiousity etc",
                                        emailInstructionsStart: '',
                                        emailOutro: 'We look forward to learning with you.'}

                         };

  var emailSignature =  "<p>Until then stay safe, sane, and sanitised, <br>" +
                        "Jen<p>" +
                        "<b>Jennifer Payne</b><br>" +
                        "Scientist & Founder<p>" +
                        '<a href="mailto:jen@STEMpals.org.au">jen@STEMpals.org.au</a><br>' +
                        "+61 437 353 827<br>" +
                        '<a href="https://www.linkedin.com/company/curiosityfactory">LinkedIn</></a><br>' +
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
      eEmailInstructions = eEmailInstructions + '<br><br><a href="' + emailDetails[e]['folderLink'] + '">Folder for your school: ' + emailDetails[e]['schoolName'] + '</a><br><br>';
      eEmailInstructions = eEmailInstructions + 'Your folder contains the following important files:<ul>'
      var welcomePackDetails = getWelcomePackContents(emailDetails[e]['schoolName'], emailVariables);
      for ( wpd of welcomePackDetails ){
        eEmailInstructions = eEmailInstructions + '<li><b>' + wpd['filename'] + '</b><br>' + wpd['fileInformation'] + ' </li>'
      }
      eEmailInstructions = eEmailInstructions + '</ul><br>';
    }

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


function getWelcomePackContents(schoolName, emailVariables){
  var welcomePackContents = [
                              {'filename': schoolName + ' Student list',
                               'fileInformation': 'Please fill in this file before ' + emailVariables['studentListDueDate']},
                              {'filename': emailVariables['year'] + ' Welcome Pack',
                               'fileInformation': 'Something exciting about the welcome pack'},
                            ];
  return welcomePackContents;
}