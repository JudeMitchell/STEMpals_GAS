// function getSchoolFolders() {
function getSchoolFolders(yearFolderID, subfolder, trimToSchoolNameOnly=false) {
  var yearFolder = DriveApp.getFolderById(yearFolderID);
  var yearSchoolsFolder = yearFolder.getFoldersByName('Schools').next();
  var schoolFolders = yearSchoolsFolder.getFolders();

  var folderDict = {};
  while ( schoolFolders.hasNext() ){
    var sf = schoolFolders.next();
    if ( sf.getName() == 'Pairing Surveys'){
      continue;
    }
    if ( subfolder == 'main' ){
      if ( trimToSchoolNameOnly ){
        var sfKey = sf.getName().slice(11);
      } else {
        var sfKey = sf.getName();
      }
      folderDict[sfKey] = sf;
    }
  }

  return folderDict;
}
