function resetSheetSize(sheet, rowsRequired, columnsRequired){
  // DOES NOT CLEAR DATA
  if ( sheet.getFilter() != null ){  
    sheet.getFilter().remove();
  }

  var frozenRows = sheet.getFrozenRows();
  if ( frozenRows > 0 ){
    sheet.setFrozenRows(0);
  }

  var currentRows = sheet.getMaxRows();
  var rowsDiff = rowsRequired - currentRows;

  if ( rowsDiff > 0 ){
    sheet.insertRowsAfter(currentRows, rowsDiff);
  } else if ( rowsDiff < 0 ){
    sheet.deleteRows(sheet.getMaxRows() + rowsDiff, currentRows - rowsRequired);
  }

  var currentColumns = sheet.getMaxColumns();
  var clmnsDiff = columnsRequired - currentColumns;
  
  if ( clmnsDiff > 0 ){
    sheet.insertColumnsAfter(currentColumns, clmnsDiff);
  } else if ( clmnsDiff < 0 ){
    sheet.deleteColumns(sheet.getMaxColumns() + clmnsDiff + 1, currentColumns - columnsRequired);
  }

  if ( frozenRows > 0 & rowsRequired > frozenRows ){
    sheet.setFrozenRows(frozenRows);
  }

}

function resetSheetData(sheet, firstDataRow, firstDataCol, dataSet) {
  // clears everything right and below of firstDataRow, firstDataCol, and replaces with dataSet
  // removes or adds rows and columns as required
  // if ( sheet.getFilter() != null ){  
  //   sheet.getFilter().remove();
  // }

  var frozenRows = sheet.getFrozenRows();
  // if ( frozenRows > 0 ){
  //   sheet.setFrozenRows(0);
  // }
  
  // var currentRows = sheet.getMaxRows();
  var dataRowsRequired = dataSet.length;
  var rowsRequired = dataRowsRequired + firstDataRow - 1;

  // if ( currentRows < rowsRequired ){
  //   sheet.insertRowsAfter(currentRows, rowsRequired - currentRows);
  // } else if ( rowsRequired < currentRows ){
  //   sheet.deleteRows(firstDataRow, currentRows - rowsRequired);
  // }

  var currentColumns = sheet.getMaxColumns();
  if ( dataSet.length > 0 ){
    var dataColumnsRequired = dataSet[0].length;
    var columnsRequired = dataColumnsRequired + firstDataCol - 1;
  } else {
    var columnsRequired = currentColumns;
  }

  resetSheetSize(sheet, rowsRequired, columnsRequired);

  // if ( currentColumns < columnsRequired ){
  //   sheet.insertColumnsAfter(currentColumns, columnsRequired - currentColumns);
  // } else if ( columnsRequired < currentColumns ){
  //   sheet.deleteColumns(firstDataCol, currentColumns - columnsRequired);
  // }

  if ( dataRowsRequired > 0 ){
    sheet.getRange(firstDataRow, firstDataCol, sheet.getMaxRows() - firstDataRow + 1, sheet.getMaxColumns() - firstDataCol + 1).clear();
    sheet.getRange(firstDataRow, firstDataCol, dataSet.length, dataSet[0].length).setValues(dataSet);
  }

  if ( frozenRows > 0 & dataRowsRequired > 0 ){
    sheet.setFrozenRows(frozenRows);
  }
}


function updateSheetDataInPlace(sheet, firstDataRow, firstDataClmn, dataSet, idClmnIndex, expectedHeaders, alerts){
  if ( sheet.getFilter() != null ){  
    sheet.getFilter().remove();
  }

  var headerRowNum = firstDataRow - 1;
  var sheetHeaders = sheet.getRange(headerRowNum, firstDataClmn, 1, expectedHeaders[0].length).getValues();

  if ( expectedHeaders[0].join('|') != sheetHeaders[0].join('|') ){
    addAlert(alerts, 'sheetHeaderMismatch', sheet.getName(), sheet.getParent().getName());
    return;
  }

  var idClmnNbr = idClmnIndex + 1;
  var sheetIdClmn = getListFromRange(sheet.getRange(firstDataRow, idClmnNbr, sheet.getMaxRows() - firstDataRow + 1, 1), 'column');
  var idsToKeep = [];

  for ( ds of dataSet ){
    var dsID = ds[idClmnIndex];
    if ( sheetIdClmn.includes(dsID) ){
      var dsRow = sheetIdClmn.indexOf(dsID) + firstDataRow;
    } else {
      sheet.insertRowAfter(sheet.getMaxRows());
      var dsRow = sheet.getMaxRows();
    }
    sheet.getRange(dsRow, firstDataClmn, 1, expectedHeaders[0].length).setValues([ds]);
    idsToKeep.push(dsID);
  }
  
  var sheetRowsToDelete = [];
  for ( sheetRow = firstDataRow; sheetRow <= sheet.getMaxRows(); sheetRow++ ){
    var sheetRowID = sheet.getRange(sheetRow, idClmnNbr, 1, 1).getValue();
    if ( ! idsToKeep.includes(sheetRowID) ){
      sheetRowsToDelete.push(sheetRow);
    }
  }
  deleteListOfRows(sheet, sheetRowsToDelete);
  
  return alerts;
}