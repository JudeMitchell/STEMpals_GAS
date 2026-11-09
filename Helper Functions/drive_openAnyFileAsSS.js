function openAnyFileAsSS(fileObject, folderObject, fileType) {
  if ( fileObject.getMimeType() != MimeType.GOOGLE_SHEETS ){
    var fileBlob = fileObject.getBlob();
    const config = {
                    name: "[Temp] " + fileType,
                    mimeType: MimeType.GOOGLE_SHEETS,
                    parents: [folderObject.getId()]
                  };
    var fileConverted = Drive.Files.create(config, fileBlob);
    var ss = SpreadsheetApp.openById(fileConverted.id);
  } else {
    var ss = SpreadsheetApp.open(fileObject);
  } 
  return ss;
}
