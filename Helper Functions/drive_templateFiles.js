function copyTemplate(fs, formattedFileName, folder, textVariables){
  const fsTemplate = getTemplateFile(fs, 'file');
  var copiedFile = fsTemplate.makeCopy(formattedFileName, folder);

  if ( fs == 'PermissionSlip'){
    var copiedDoc = DocumentApp.openById(copiedFile.getId());
    var docBody = copiedDoc.getActiveTab().asDocumentTab().getBody();
    docBody.replaceText('<YYYY>', textVariables['year']);
  } else {
    Logger.log(fs)
  }
}