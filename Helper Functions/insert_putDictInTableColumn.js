function putDictInTableColumn(dataDict, sheet, keyColumn, valueColumn, firstDataRow){
  var keyColumnValues = getListFromRange(sheet.getRange(keyColumn + firstDataRow + ':' + keyColumn), 'column');
  
  for ( key of Object.keys(dataDict) ){
    if ( keyColumnValues.includes(key) ){
      var keyDataRow = keyColumnValues.indexOf(key) + firstDataRow;
    } else {
      sheet.insertRowAfter(sheet.getMaxRows());
      var keyDataRow = sheet.getMaxRows();
      sheet.getRange(keyColumn + keyDataRow).setValue(key);
    }
    sheet.getRange(valueColumn + keyDataRow).setValue(dataDict[key]);
  }
}