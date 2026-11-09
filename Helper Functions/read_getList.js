function getListFromRange(range, rangeType, includeBlanks=false) {
  var rangeValues = range.getValues();

  var rangeList = [];

  if ( rangeType === 'column' ) {
    if ( includeBlanks ) {
      for ( i = 0; i < rangeValues.length; i++ ) {
        rangeList.push(rangeValues[i][0]);
      }
    } else {
      for ( i = 0; i < rangeValues.length; i++ ) {
        if ( rangeValues[i][0] === '' ) {
          break;
        } else {
          rangeList.push(rangeValues[i][0]);
        }
      }
    }
  } else if ( rangeType === 'row' ) {
    if ( includeBlanks ) {
      for ( i = 0; i < rangeValues[0].length; i++ ) {
        rangeList.push(rangeValues[0][i]);
      }
    } else {
      for ( i = 0; i < rangeValues[0].length; i++ ) {
        if ( rangeValues[0][i] === '' ) {
          break;
        } else {
          rangeList.push(rangeValues[0][i]);
        }
      }
    }
  } else {
    Browser.msgBox('Called getListFromRange with incorrect rangeType: ' + rangeType);
    return;
  }

  return rangeList;
}

function getDistinctListFromRange(range, rangeType, includeBlanks=false) {
  var rangeValues = range.getValues();

  var rangeList = [];

  if ( rangeType === 'column' ) {
    if ( includeBlanks ) {
      for ( i = 0; i < rangeValues.length; i++ ) {
        if ( ! rangeList.includes(rangeValues[i][0]) ){
          rangeList.push(rangeValues[i][0]);
        }
      }
    } else {
      for ( i = 0; i < rangeValues.length; i++ ) {
        if ( rangeValues[i][0] === '' ) {
          break;
        } else {
          if ( ! rangeList.includes(rangeValues[i][0]) ){
            rangeList.push(rangeValues[i][0]);
          }
        }
      }
    }
  } else if ( rangeType === 'row' ) {
    if ( includeBlanks ) {
      for ( i = 0; i < rangeValues[0].length; i++ ) {
        if ( ! rangeList.includes(rangeValues[0][i]) ){
          rangeList.push(rangeValues[0][i]);
        }
      }
    } else {
      for ( i = 0; i < rangeValues[0].length; i++ ) {
        if ( rangeValues[0][i] === '' ) {
          break;
        } else {
          if ( ! rangeList.includes(rangeValues[0][i]) ){
            rangeList.push(rangeValues[0][i]);
          }
        }
      }
    }
  } else {
    Browser.msgBox('Called getListFromRange with incorrect rangeType: ' + rangeType);
    return;
  }

  return rangeList;
}
