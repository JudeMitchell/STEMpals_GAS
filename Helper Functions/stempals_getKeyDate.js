function getRunSheet(year){
  const fileManagerSS = SpreadsheetApp.openById('1zO0JTbchfqLHyYTG6PYleFsGzSrOByOphvwqThV3Fwg');
  const runSheetIDsSheet = fileManagerSS.getSheetByName('Run Sheet IDs');
  
  if ( year == null ){
    var rowNbr = 2;
  }

  var runSheetID = runSheetIDsSheet.getRange('B' + rowNbr).getValue();
  return SpreadsheetApp.openById(runSheetID);
}

function getKeyDate(dateID, format) {
  var dateID = 'Student List Due';
  var format = '<EEEE> <d> <mmmm>';
  
  const runSheetSS = getRunSheet();
  const keyDatesSheet = runSheetSS.getSheetByName('Key Dates');
  
  var dateRowFinder = keyDatesSheet.getRange('A1:A').createTextFinder(dateID).findAll();

  if ( dateRowFinder.length == 0 ){
    throw new Error('*' + dateID + '* not found in Key Dates tab!');
  } else {
    var keyDate = keyDatesSheet.getRange('B' + dateRowFinder[0].getRow()).getValue();
    if ( keyDate == '' ){
      throw new Error('*' + dateID + '* has no date entered in Key Dates tab!');
    }
    if ( format == '' ){
      return keyDate;
    } else {
      var config = {
                    'EEEE': keyDate.toLocaleString(undefined, {'weekday': 'long'}),
                    'd': keyDate.toLocaleString(undefined, {'day': 'numeric'}),
                    'mmmm': keyDate.toLocaleString(undefined, {'month': 'long'})
                    };
      for ( c of Object.keys(config) ){
        format = format.replace('<' + c + '>', config[c]);
      }
      
      return format;
    }
  }
}
