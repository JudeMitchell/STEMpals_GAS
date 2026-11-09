function addListToColumn(list, sheet, rowNumber, colNumber) {
  var column = [];

  for ( li of list ){
    column.push([li]);
  }

  sheet.getRange(rowNumber, colNumber, column.length, 1).setValues(column);
}
