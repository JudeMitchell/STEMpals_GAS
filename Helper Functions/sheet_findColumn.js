function findColumnNumber(sheet, columnName, headerSearchColumn, headerSearchValue) {
  return findColumn('number', sheet, columnName, headerSearchColumn, headerSearchValue)
}

function findColumnIndex(sheet, columnName, headerSearchColumn, headerSearchValue) {
  return findColumn('index', sheet, columnName, headerSearchColumn, headerSearchValue)
}

function findColumnLetter(sheet, columnName, headerSearchColumn, headerSearchValue) {
  return findColumn('letter', sheet, columnName, headerSearchColumn, headerSearchValue)
}

function findColumn(returnValue, sheet, columnName, headerSearchColumn, headerSearchValue) {
  // var headerRow = findFirstDataRow(sheet, headerSearchColumn, headerSearchValue) - 1;
  
  // var iCell = sheet.getRange('A' + headerRow);

  var tf = sheet.createTextFinder(columnName).matchEntireCell(true).matchCase(true);
  if ( tf.findAll().length == 1 ){
    var iCell = tf.findAll()[0];
  } else if ( tf.findAll().length == 0 ){
    var helperError = new Error("InputError");
    helperError.isCustom = true;
    helperError.message = "No cell contains column name (" + columnName + ") in " + sheet.getName();
    Logger.log(helperError.message);
    throw helperError;
  } else {
    var helperError = new Error("InputError");
    helperError.isCustom = true;
    helperError.message = "Multiple (" + tf.findAll().length + ") cells contains column name (" + columnName + ") in " + sheet.getName();
    Logger.log(helperError.message);
    throw helperError;
  }

  
  // for ( colJumper = 0; colJumper < sheet.getMaxColumns(); colJumper ++ ){
    // if ( iCell.getValue() == columnName ){
  if ( returnValue == 'number' ){
    return iCell.getColumn();
  } else if ( returnValue == 'index' ){
    return iCell.getColumn() - 1;
  } else if ( returnValue == 'letter' ){
    return iCell.getA1Notation().match(/([A-Z]*)([0-9]*)/)[1];
  }
    // } else {
    //   var iCell = iCell.offset(0, 1);
    // }
  // }

  var helperError = new Error("InputError")
  helperError.isCustom = true
  helperError.message = "Cannot find column (" + columnName + ") in header row of " + sheet.getName() + " (looking in row " + headerRow + ")";
  throw helperError
}

function findAllColumns(returnValue, sheet, exampleColumnName, stopAtBlank=true){
  var clmnDict = {};
  var tf = sheet.createTextFinder(exampleColumnName).matchEntireCell(true).matchCase(true);
  if ( tf.findAll().length == 1 ){
    var headerRowNum = tf.findAll()[0].getRow();
    for ( headerClmn = 1; headerClmn <= sheet.getMaxColumns(); headerClmn++ ){
      var iClmnCell = sheet.getRange(headerRowNum, headerClmn, 1, 1)
      var iClmnName = iClmnCell.getValue();

      if ( stopAtBlank & iClmnName == '' ){
        return clmnDict;
      } else if (iClmnName == '' ){
        continue;
      }

      if ( returnValue == 'letter' ){
        clmnDict[iClmnName] = iClmnCell.getA1Notation().replace(headerRowNum, '');
      } else if ( returnValue == 'number' ) {
        clmnDict[iClmnName] = iClmnCell.getColumn();
      } else if ( returnValue == 'index' ){
        clmnDict[iClmnName] = iClmnCell.getColumn() - 1;
      } else {
        var helperError = new Error("InputError");
        helperError.isCustom = true;
        helperError.message = "Cannot return header column detail: " + returnValue;
        Logger.log(helperError.message);
        throw helperError;
      }
    }

  } else if ( tf.findAll().length == 0 ){
    var helperError = new Error("InputError");
    helperError.isCustom = true;
    helperError.message = "No cell contains column name (" + exampleColumnName + ") in " + sheet.getName();
    Logger.log(helperError.message);
    throw helperError;
  } else {
    var helperError = new Error("InputError");
    helperError.isCustom = true;
    helperError.message = "Multiple (" + tf.findAll().length + ") cells contains column name (" + exampleColumnName + ") in " + sheet.getName();
    Logger.log(helperError.message);
    throw helperError;
  }

  return clmnDict;
}