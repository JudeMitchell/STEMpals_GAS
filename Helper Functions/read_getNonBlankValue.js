function getNonBlankValue(cellRange, fieldName) {
  var value = cellRange.getValue();
  if ( value == '' ){
    // cellRange.setBackground('pink');
    Logger.log(cellRange.getA1Notation())
    throw new Error('Empty field in row ' + cellRange.getRow() + ': ' + fieldName);
  }
  return value;
}
