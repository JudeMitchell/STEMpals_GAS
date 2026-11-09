/**
 * Converts a spreadsheet range into a JavaScript dictionary (object), mapping a key column index 
 * (or pair of column indices) to values from a specified value column index, with options to group into arrays.
 *
 * @param {GoogleAppsScript.Spreadsheet.Range} range - The spreadsheet range containing the data.
 * @param {number|Array<number>} keyColIndex - The zero-based index of the key column, or an array of two indices to combine.
 * @param {number} valueColIndex - The zero-based index of the value column.
 * @param {boolean} [allLists=false] - Whether to store single values in an array instead of just the value.
 * @param {boolean} [distinctLists=false] - Whether to keep only distinct/unique values in the arrays when `allLists` is true.
 * @return {Object.<string, any>} The constructed dictionary mapping keys to values or lists of values.
 * @throws {Error} Throws an `InputError` if `keyColIndex` is an array with more than 2 elements.
 */
function getDictFromRange(range, keyColIndex, valueColIndex, allLists=false, distinctLists=false) {
  // var nesting = [{'name': 'Sub Region',
  //                 'colIndex': 2,
  //                 'levelType': 'unique'}]
  
  // currently works with assigning one key value level
  // keyColIndex can take a list of two items, the first and last indexes of the columns to include
  var dict = {};
  var data = range.getValues();

  for ( dp of data ){
    if ( typeof(keyColIndex) == 'number' ){
      var dpKey = dp[keyColIndex];
      if ( dp[keyColIndex] == '' ){
        break;
      }
    } else {
      if ( keyColIndex.length > 2 ){
        var helperError = new Error("InputError");
        helperError.isCustom = true;
        helperError.message = "getDictFromRange() not prepared for a list of greater than 2 for keyColIndex";
        throw helperError;
      }
      // var dpKey = dp.slice(keyColIndex[0], keyColIndex[1] + 1).join(" ");
      
      if ( dp[keyColIndex[0]] == '' ){
        break;
      }
      var dpKey = dp[keyColIndex[0]] + ' ' + dp[keyColIndex[1]];
    }

    if ( typeof(dpKey) == 'number' ){
      dpKey = dpKey.toString();
    }

    if ( allLists ){
      if ( ! Object.keys(dict).includes(dpKey) ){
        dict[dpKey] = [];
      }
      if ( distinctLists ){
        if ( !dict[dpKey].includes(dp[valueColIndex]) ){
          dict[dpKey].push(dp[valueColIndex]);
        }
      } else {
        dict[dpKey].push(dp[valueColIndex]);
      }
      
    } else {
      if ( Object.keys(dict).includes(dpKey) ){
        if ( Array.isArray(dict[dpKey]) ){
          dict[dpKey].push(dp[valueColIndex]);
        } else {
          dict[dpKey] = [dict[dpKey], dp[valueColIndex]];
        }
      } else {
        dict[dpKey] = dp[valueColIndex];
      }
    }
  }
  return dict;
}

/**
 * Builds a dictionary from spreadsheet rows where each key maps to an array of all the
 * non-blank values collected from multiple specified columns.
 *
 * @param {GoogleAppsScript.Spreadsheet.Range} range - The spreadsheet range containing the rows.
 * @param {number} keyColIndex - The zero-based index of the column to use as dictionary keys.
 * @param {Array<number>} valueColIndexList - An array of zero-based column indices whose non-blank values form the list for each key.
 * @return {Object.<string, Array<any>>} A dictionary mapping each key to an array of extracted row values.
 */
function getDictofListsFromRows(range, keyColIndex, valueColIndexList){
  var values = range.getValues();
  var returnDict = {};
  for ( v of values ){
    var vList = [];
    for ( vci of valueColIndexList ){
      if ( v[vci] != '' ){
        vList.push(v[vci]); 
      }
    }
    returnDict[v[keyColIndex]] = vList;
  }
  return returnDict;
}

/**
 * Converts a spreadsheet range into a dictionary based on key and value columns, 
 * filtering rows against a set of required conditions before processing.
 *
 * @param {GoogleAppsScript.Spreadsheet.Range} range - The spreadsheet range containing the data.
 * @param {number|Array<number>} keyColIndex - The zero-based index or 2-item index pair for the key column.
 * @param {number} valueColIndex - The zero-based index of the value column.
 * @param {Array<Object>} conditions - An array of condition dicts defining column indices (condClmnInd) and an array of allowed/kept values (keepValues).
 * @param {boolean} [allLists=false] - Whether to store all matching values in arrays per key.
 * @param {boolean} [distinctLists=false] - Whether to store only distinct/unique values in arrays when `allLists` is true.
 * @param {string} stopAtBlank - Whether to stop reading rows where *key* is blank or where *value* is blank.
 * @return {Object.<string, any>} The filtered and constructed dictionary.
 * @throws {Error} Throws an `InputError` if `keyColIndex` contains a list greater than 2 elements.
 */
function getDictFromRangeIf(range, keyColIndex, valueColIndex, conditions, allLists=false, distinctLists=false, stopAtBlank="key") {
  // [{'condClmnInd': changeTabClmnInds['Set inactive?'], 'keepValues': [true]}];

  // currently works with assigning one key value level
  // keyColIndex can take a list of two items, the first and last indexes of the columns to include
  var dict = {};
  var data = range.getValues();

  dpLoop: for ( dp of data ){
    if ( stopAtBlank == 'value' ){
      if ( dp[valueColIndex] == '' ){
        break;
      }
    } else {
      if ( typeof(keyColIndex) == 'number' ){
        if ( dp[keyColIndex] == '' ){
          break;
        }
      } else {
        if ( dp[keyColIndex[0]] == '' ){
          break;
        }
      }
    }

    for ( cnd of conditions ){
      // Logger.log(cnd);
      // Logger.log(dp[cnd['condClmnInd']]);
      if ( ! cnd['keepValues'].includes(dp[cnd['condClmnInd']]) ){
        continue dpLoop;
      }
    }

    if ( typeof(keyColIndex) == 'number' ){
      var dpKey = dp[keyColIndex];
    } else {
      if ( keyColIndex.length > 2 ){
        var helperError = new Error("InputError");
        helperError.isCustom = true;
        helperError.message = "getDictFromRange() not prepared for a list of greater than 2 for keyColIndex";
        throw helperError;
      }
      // var dpKey = dp.slice(keyColIndex[0], keyColIndex[1] + 1).join(" ");
      var dpKey = dp[keyColIndex[0]] + ' ' + dp[keyColIndex[1]];
    }

    if ( typeof(dpKey) == 'number' ){
      dpKey = dpKey.toString();
    }

    if ( allLists ){
      if ( ! Object.keys(dict).includes(dpKey) ){
        dict[dpKey] = [];
      }
      if ( distinctLists ){
        if ( !dict[dpKey].includes(dp[valueColIndex]) ){
          dict[dpKey].push(dp[valueColIndex]);
        }
      } else {
        dict[dpKey].push(dp[valueColIndex]);
      }
      
    } else {
      if ( Object.keys(dict).includes(dpKey) ){
        if ( Array.isArray(dict[dpKey]) ){
          dict[dpKey].push(dp[valueColIndex]);
        } else {
          dict[dpKey] = [dict[dpKey], dp[valueColIndex]];
        }
      } else {
        dict[dpKey] = dp[valueColIndex];
      }
    }
  }
  return dict;
}

/**
 * Extracts specified columns from a table of data into an array of dictionaries, 
 * where each dict contains the requested column names as keys, with values from a single row of data.
 *
 * @param {number} headerRow - The row number where table headers are located.
 * @param {Array<string>} columnNames - An array of column header names to extract.
 * @param {number} nDataRows - The number of data rows to process below the header.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The target Google Sheet tab.
 * @param {Array<Object>} An array of dicts representing the rows.
 */
function getDictListsFromTable(headerRow, columnNames, nDataRows, sheet){
  var firstDataRow = headerRow + 1;
  var headerIndex = getListFromRange(sheet.getRange(headerRow, 1, 1, sheet.getMaxColumns()), 'row');

  if ( nDataRows == null ){
    var nDataRows = sheet.getMaxRows() - headerRow;
  }

  var listOutput = [];

  for ( ndr = 1; ndr <= nDataRows; ndr++ ){
    listOutput.push({});
  }
  Logger.log(headerIndex)
  for ( cn of columnNames ){
    var cnNum = headerIndex.indexOf(cn) + 1;
    for ( dictNbr = 0; dictNbr < nDataRows; dictNbr++ ){
      listOutput[dictNbr][cn] = sheet.getRange(dictNbr + firstDataRow, cnNum, 1, 1).getValue();
    }
  }
  return listOutput;
}

/**
 * Extracts specified columns from a table into a nested dictionary, 
 * keyed first by a primary key column value and then by column names.
 *
 * @param {number} headerRow - The row number where table headers are located.
 * @param {string} keyColumnName - The header name of the column used as primary keys.
 * @param {Array<string>} columnNames - An array of column header names to include in the inner dictionaries.
 * @param {number} nDataRows - The number of data rows to process.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The target Google Sheet tab.
 * @return {Object.<string, Object.<string, any>>} A nested dictionary structure of the table data.
 */
function getDictDictsFromTable(headerRow, keyColumnName, columnNames, nDataRows, sheet){
  var firstDataRow = headerRow + 1;
  var headerIndex = getListFromRange(sheet.getRange(headerRow, 1, 1, sheet.getMaxColumns()), 'row', true);
  if ( nDataRows == null ){
    var nDataRows = sheet.getMaxRows() - headerRow;
  }

  var dictOutput = {};
  var keyClmnNum = headerIndex.indexOf(keyColumnName) + 1;

  for ( ndr = 0; ndr < nDataRows; ndr++ ){
    var ndrKey = sheet.getRange(ndr + firstDataRow, keyClmnNum, 1, 1).getValue();
    dictOutput[ndrKey] = {};
    for ( cn of columnNames ){
      var cnNum = headerIndex.indexOf(cn) + 1;
      dictOutput[ndrKey][cn] = sheet.getRange(ndr + firstDataRow, cnNum, 1, 1).getValue();
    }
  }
  return dictOutput;
}
