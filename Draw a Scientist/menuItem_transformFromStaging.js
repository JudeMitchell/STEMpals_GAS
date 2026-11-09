function transformFromStaging() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const stagingTab = ss.getSheetByName('data staging').activate();
  const wideStagingTab = ss.getSheetByName('wide data staging').activate();

  // stagingTab.insertColumnAfter(2);
  // stagingTab.getRange('C1').setValue('Class');
  // stagingTab.insertColumnsAfter(4,2);
  // stagingTab.getRange('E1').setValue('student ID');
  // stagingTab.getRange('F1').setValue('Term');

  // var fileNames = HelperFunctions.getListFromRange(stagingTab.getRange('A2:A'), 'column');
  // var fileNameRegex = /([A-z|0-9]+)_[a-z]+([1-4])-.+/;

  // for ( i=0; i<fileNames.length; i++ ){
  //   var match = fileNames[i].match(fileNameRegex);
  //   stagingTab.getRange('C' + Number(i + 2)).setValue(match[1]);
  //   stagingTab.getRange('E' + Number(i + 2)).setValue(match[1].toUpperCase() + '-' + stagingTab.getRange('D' + Number(i + 2)).getValue().toUpperCase());
  //   stagingTab.getRange('F' + Number(i + 2)).setValue(match[2]);
  // }



  var rawData = stagingTab.getRange('E2:Q' + stagingTab.getMaxRows()).getValues();
  //TODO make a thing for looking at comments

  // based on range starting at E
  var studentIndex = 0;
  var termIndex = 1;
  var numLocationIndex = 2;
  var locationDesc = 3;
  var numPeopleIndex = 5;
  var stemProDescIndex = 6;

  var studentsInWideStaging = [];

  for ( const datapoint of rawData ){
    if ( studentsInWideStaging.includes(datapoint[studentIndex]) ){
      var studentRow = studentsInWideStaging.indexOf(datapoint[studentIndex]) + 3;
    } else {
      var studentRow = wideStagingTab.getMaxRows(); //sheet should start with a blank row
      wideStagingTab.insertRowAfter(wideStagingTab.getMaxRows());
      studentsInWideStaging.push(datapoint[studentIndex]);
    }

    if ( datapoint[termIndex] === 1 ) {
      var wideDataCols = {'drawingFound': 'B',
                          'numLocations': 'C',
                          'numPeople': 'D',
                          'stemProGender': 'E',
                          'stemProEthnicity': 'F',
                          'stemProHumanised': 'G',
                          'stemType': 'H',
                          'stemProNamed': 'I',
                          };
    } else {
      var wideDataCols = {'drawingFound': 'J',
                          'numLocations': 'K',
                          'numPeople': 'L',
                          'stemProGender': 'M',
                          'stemProEthnicity': 'N',
                          'stemProHumanised': 'O',
                          'stemType': 'P',
                          'stemProNamed': 'Q',
                          };
    }

    wideStagingTab.getRange('A' + studentRow).setValue([datapoint[studentIndex]]);
    wideStagingTab.getRange(wideDataCols['drawingFound'] + studentRow).setValue(true);
    wideStagingTab.getRange(wideDataCols['numPeople'] + studentRow).setValue(datapoint[numPeopleIndex]);
    
    // gender from 2024 format
    var datapointGenderString = '';
    if ( /A/.test(datapoint[stemProDescIndex]) ){
      datapointGenderString = datapointGenderString + 'female, ';
    }
    if ( /B/.test(datapoint[stemProDescIndex]) ){
      datapointGenderString = datapointGenderString + 'male, ';
    }
    if ( /C/.test(datapoint[stemProDescIndex]) ){
      datapointGenderString = datapointGenderString + 'non binary, ';
    }
    if ( datapointGenderString === '' ){
      datapointGenderString = 'not determined, ';
    }
    wideStagingTab.getRange(wideDataCols['stemProGender'] + studentRow).setValue(datapointGenderString.slice(0,-2));

    // ethnicity from 2024 format
    var datapointEthnicityString = '';
    if ( /E/.test(datapoint[stemProDescIndex]) ){
      datapointEthnicityString = datapointEthnicityString + 'white, ';
    }
    if ( /D/.test(datapoint[stemProDescIndex]) ){
      datapointEthnicityString = datapointEthnicityString + 'minority, ';
    }
    if ( datapointGenderString === '' ){
      datapointEthnicityString = 'unknown, ';
    }
    wideStagingTab.getRange(wideDataCols['stemProEthnicity'] + studentRow).setValue(datapointEthnicityString.slice(0,-2));

  }
  
}
