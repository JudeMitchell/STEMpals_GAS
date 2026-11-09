function getFolderStructure(folderRef) {
  const fileManagerSS = SpreadsheetApp.openById('1zO0JTbchfqLHyYTG6PYleFsGzSrOByOphvwqThV3Fwg');
  const folderStructureSheet = fileManagerSS.getSheetByName('Folder structure');

  var folderstructureClmnNbrs = findAllColumns('number', folderStructureSheet, 'YYYY STEMpals Program', false);
  var nRows = getListFromRange(folderStructureSheet.getRange(2, folderstructureClmnNbrs[folderRef], folderStructureSheet.getMaxRows() - 1, 1), 'column').length;
  
  // Logger.log(folderstructureClmnNbrs[folderRef])
  var detailClmnFormulas = folderStructureSheet.getRange(1, folderstructureClmnNbrs[folderRef] + 1, 1, 2).getFormulas();
  folderStructureSheet.getRange(1, folderstructureClmnNbrs[folderRef] + 1, 1, 2).setValues([['Type', 'Create at']]);
  // var columnNames = [ 'Type (' + folderRef + ')',
  //                     'Create at (' + folderRef + ')'
  //                   ];
  var folderDataRange = folderStructureSheet.getRange(2, folderstructureClmnNbrs[folderRef], folderStructureSheet.getMaxRows(), 3);
  // Logger.log(folderDataRange.getA1Notation())
  // var structure = getDictofListsFromRows(folderDataRange, )
  // var structure = getDictListsFromTable(firstRow, ['Object Name', 'Type', 'Create at'], nRows, folderStructureSheet)
  var structure = getDictDictsFromTable(1, folderRef, ['Type', 'Create at'], nRows, folderStructureSheet);
  // Logger.log(structure)

  folderStructureSheet.getRange(1, folderstructureClmnNbrs[folderRef] + 1, 1, 2).setFormulas(detailClmnFormulas);
  return structure;
}

function updateSubFolder(subFolder, folderRef, yearFolder){
  var folderType = 'Schools';
  var folderRef = 'STEMpalsYY_<SchoolName>';
  var schoolFolder = DriveApp.getFolderById('1Tdi82prS6o5MZ3zZYLfEOOyoYbAs-7Ir')


  updateFolder(folderRef, schoolFolder, '2027', 'WelcomePack');

}

function getFileTreatment(){
  var fileTreatment = { 'PDF': 'blobCopy',
                        'docx': 'blobCopy',
                        'Google Sheet': 'driveCopy'};
  return fileTreatment;
}

function updateFolder(folderRef, folder, year, timePoint){
  var yearResourcesFolder = getYearFolder(year, 'Teacher Resources');

  if ( folderRef == 'STEMpalsYY_<SchoolName>' ){
    var schoolName = folder.getName().split('_')[1];
  } else {
    var schoolName = '';
  }

  var folderStructure = getFolderStructure(folderRef);
  var fileTreatment = getFileTreatment();

  var folderFolderItr = folder.getFolders();
  var folderFolderNames = [];
  while ( folderFolderItr.hasNext() ){
    var fFolder = folderFolderItr.next();
    folderFolderNames.push(fFolder.getName());
  }

  var folderFileItr = folder.getFiles();
  var folderFileNames = [];
  while ( folderFileItr.hasNext() ){
    var fFile = folderFileItr.next();
    folderFileNames.push(fFile.getName());
  }

  for ( fs of Object.keys(folderStructure) ){
    if ( folderStructure[fs]['Create at'] == timePoint ){
      var formattedFolderName = fs.replace('YYYY', year).replace('YY', year.slice(2,4));
      if ( folderStructure[fs]['Type'] == 'Folder' ){
        if ( folderFolderNames.includes(formattedFolderName) ){
          Logger.log('  folder already exists: ' + formattedFolderName);
        } else {
          folder.createFolder(formattedFolderName);
        }
      } else {
        var fsFileType = folderStructure[fs]['Type'];
        if ( ! /Google/.test(fsFileType) ){
          var formattedFileName = fs.replace('YYYY', year).replace('YY', year.slice(2,4)).replace('<SchoolName>', schoolName) + '.' + fsFileType.toLowerCase();
        } else {
          var formattedFileName = fs.replace('YYYY', year).replace('YY', year.slice(2,4)).replace('<SchoolName>', schoolName);
        }
        if ( folderFileNames.includes(formattedFileName) ){
          Logger.log('  file already exists: ' + formattedFileName);
        } else {          
          if ( fileTreatment[fsFileType] == 'blobCopy' ){
            var fileToCopy = yearResourcesFolder.getFilesByName(formattedFileName).next();
            fileToCopy.makeCopy(formattedFileName, folder);
          } else if ( fileTreatment[fsFileType] == 'driveCopy' ) {
            const fsTemplate = getTemplateFile(fs, 'file');
            fsTemplate.makeCopy(formattedFileName, folder);


          }
        }
      }
    }
  }
}

function getYearFolder(year, subfolder){
  const fileManagerSS = SpreadsheetApp.openById('1zO0JTbchfqLHyYTG6PYleFsGzSrOByOphvwqThV3Fwg');
  const folderIDsSheet = fileManagerSS.getSheetByName('Folder IDs');
  
  if ( year == null ){
    var rowNbr = 2;
  } else {
    var yearRowFinder = folderIDsSheet.getRange('A2:A').createTextFinder(year).findAll();
    var rowNbr = yearRowFinder[0].getRow();
  }
  
  var folderID = folderIDsSheet.getRange('B' + rowNbr).getValue();
  if ( subfolder == null ){
    return DriveApp.getFolderById(folderID);
  } else {
    var yearFolder = DriveApp.getFolderById(folderID);
    return yearFolder.getFoldersByName(subfolder).next()
  }
}

function getTemplateFile(fileName, returnObject){
  const controlPanel = SpreadsheetApp.openById('1zO0JTbchfqLHyYTG6PYleFsGzSrOByOphvwqThV3Fwg');
  const templateIDSheet = controlPanel.getSheetByName('Template File IDs');

  var fileIDsDict = getDictFromRange(templateIDSheet.getRange('A2:B'), 0, 1);

  if ( ! Object.keys(fileIDsDict).includes(fileName) ){
    throw new Error ('Unknown template file name (' + fileName + '), use one of: ' + Object.keys(fileIDsDict).join(', '));
  }

  if ( returnObject == 'ss' ){
    return SpreadsheetApp.openById(fileIDsDict[fileName]);
  } else if ( returnObject == 'file' ){
    return DriveApp.getFileById(fileIDsDict[fileName]);
  } else if ( returnObject == 'id' ){
    return fileIDsDict[fileName];
  } else {
    throw new Error ('Unknown returnObject (' + returnObject + ') for template file, use one of: ss, file, id');
  }
}

function checkForInputFiles(folderRef, timePoint, year, ui){
  var folderStructure = getFolderStructure(folderRef);
  var yearResourcesFolder = getYearFolder(null, 'Teacher Resources');

  var missingFiles = [];
  for ( fs of Object.keys(folderStructure) ){
    if ( folderStructure[fs]['Type'] == 'Folder' ){
      continue;
    } else {
      if ( folderStructure[fs]['Create at'] == timePoint ){
        var fsFileType = folderStructure[fs]['Type'];
        if ( ! /Google/.test(fsFileType) ){
          var formattedFileName = fs.replace('YYYY', year).replace('YY', year.slice(2,4)) + '.' + fsFileType.toLowerCase();
          var fileToCopy = yearResourcesFolder.getFilesByName(formattedFileName);
          if ( ! fileToCopy.hasNext() ){
            missingFiles.push(formattedFileName);
          }
        } else {
          var formattedFileName = fs.replace('YYYY', year).replace('YY', year.slice(2,4));
        }
      }
    }
  }

  if ( missingFiles.length > 0 ){
    ui.alert('Files missing from Teacher Resources folder!', '- ' + missingFiles.join('\n- '), ui.ButtonSet.OK);
    throw new Error ('Files missing from Teacher Resources folder: ' + missingFiles.join(', '));
  }
  return;
}  