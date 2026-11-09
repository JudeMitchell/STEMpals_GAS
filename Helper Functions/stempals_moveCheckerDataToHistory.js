function moveCheckerDataToHistory(type, ss) {
  // keep the links in for now
  // // testing
  // var ss = SpreadsheetApp.openById('1HiARVSUmR411KvAEXY1hjlVvMKxR72XMP7lnoskbOm8');
  // var type = 'safety';
  // // testing

  var historyTab = ss.getSheetByName('Historical Data');
  var newDataTab = ss.getSheetByName('New Data');

  var newDataTabFirstRow = 2;

  if ( type == 'safety' ){
    var firstInputDataCol = 'D';
    var lastInputDataCol = 'N';
    var richTextInputColumn = 'C';
    var richTextOutputColumn = 'C';
    var firstOutputDataCol = 'E';
    var lastOutputDataCol = 'O';
    var fileNameOutputCol = 'C';
    var qrOutputDataCol = 'D';
    var noOrOKFlagsCol = 'P';
    var handWrittenOrComboCol = 'Q';
    var redactedOrNotCol = 'R';
  } else {
    Browser.msgBox('This function is not ready to handle data type of: ' + type);
  }

  var firstHistoryRow = historyTab.getMaxRows() + 1;
  var lastNewDataRow = newDataTab.getMaxRows();

  var newDataYearLetterCols = newDataTab.getRange('A' + newDataTabFirstRow + ":B" + lastNewDataRow).getValues();
  var newDataDataCols = newDataTab.getRange(firstInputDataCol + newDataTabFirstRow + ":" + lastInputDataCol + lastNewDataRow).getValues();
  var newRichTextDataCol = newDataTab.getRange(richTextInputColumn + '2:' + richTextInputColumn).getRichTextValues();

  historyTab.insertRowsAfter(historyTab.getMaxRows(), newDataYearLetterCols.length);
  var lastHistoryRow = historyTab.getMaxRows();
  historyTab.getRange('A' + firstHistoryRow + ":B" + lastHistoryRow).setValues(newDataYearLetterCols);
  historyTab.getRange(firstOutputDataCol + firstHistoryRow + ":" + lastOutputDataCol + lastHistoryRow).setValues(newDataDataCols);
  for ( i = 0; i < newRichTextDataCol.length; i++ ){
    var historyRowNum = i + firstHistoryRow;
    historyTab.getRange(richTextOutputColumn + historyRowNum).setRichTextValue(newRichTextDataCol[i][0]);
    if ( newRichTextDataCol[i][0].getLinkUrl() != null ){
      historyTab.getRange(richTextOutputColumn + historyRowNum).setFontColor('#1155cc');
    }

    var iQRMatching = historyTab.getRange(fileNameOutputCol + historyRowNum).getValue().match(/([A-z|\-]*)([0-9]*-{0,1}[0-9]+$)/);
    if ( iQRMatching != null ){
      if ( iQRMatching.length == 3 ){
        historyTab.getRange(qrOutputDataCol + historyRowNum).setValue(iQRMatching[2]);
      }
    }

    if ( type == 'safety' ){
      var noOrOKFormula = '=if(K'+ historyRowNum + '="No","Yes",IF(N'+ historyRowNum + '="OK in context","Yes","No"))';
      historyTab.getRange(noOrOKFlagsCol + historyRowNum).setValue(noOrOKFormula);
      var handWrittenOrComboFormula = '=OR(G'+ historyRowNum + '="hand written",G'+ historyRowNum + '="combination")';
      historyTab.getRange(handWrittenOrComboCol + historyRowNum).setValue(handWrittenOrComboFormula);
      var redactedOrNotFormula = '=if(N'+ historyRowNum + '="Redacted","Redacted","other")';
      historyTab.getRange(redactedOrNotCol + historyRowNum).setValue(redactedOrNotFormula);
    }
  }

  newDataTab.setFrozenRows(0);
  newDataTab.deleteRows(2, newDataTab.getMaxRows() - 1);
  newDataTab.getRange('A1:1').setFontWeight('normal');
  Browser.msgBox('All done!');
}
