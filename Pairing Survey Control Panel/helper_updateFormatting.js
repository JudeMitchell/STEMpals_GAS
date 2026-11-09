function updateFormatting(sheet) {
  var cfRules = [];
  
  if ( sheet.getName() == 'School Summary' ){
    
    var sheetMissingResponsesClmnLtr = STEMpalsHelperFunctions.findColumnLetter(sheet, 'Missing Responses');
    
    sheet.setConditionalFormatRules([]);

    // registered students CF
    var registeredRange = sheet.getRange('C2:C');
    var allRegisteredCF = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=C2=B2').setBackground('#34a853').setRanges([registeredRange]).build(); //green
    var marginOfErrorRegisteredCF = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=and(C2>=B2*0.9,C2<=B2*1.1)').setBackground('#93c47d').setRanges([registeredRange]).build(); //lighter green
    var noneRegisteredCF = SpreadsheetApp.newConditionalFormatRule().whenNumberEqualTo(0).setBackground('#ff6d01').setRanges([registeredRange]).build(); //orange
    var notEnoughRegisteredCF = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied("=C2<B2").setBackground('#fbbc04').setRanges([registeredRange]).build(); //yellow
    var noFilesCF = SpreadsheetApp.newConditionalFormatRule().whenTextContains('no files').setBackground('#ff6d01').setRanges([registeredRange]).build(); //orange
    var tooManyRegisteredCF = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied("=C2>B2").setBackground('#cc0100').setRanges([registeredRange]).build() //red
    cfRules.push(...[allRegisteredCF, marginOfErrorRegisteredCF, noneRegisteredCF, notEnoughRegisteredCF, noFilesCF, tooManyRegisteredCF]);


    // missing responses CF
    var responseCountRange = sheet.getRange(sheetMissingResponsesClmnLtr + '2:' + sheetMissingResponsesClmnLtr);
    var emptyCF = SpreadsheetApp.newConditionalFormatRule().whenCellEmpty().setBackground("white").setRanges([responseCountRange]).build();
    var allRespondedCF = SpreadsheetApp.newConditionalFormatRule().whenNumberEqualTo(0).setBackground('#34a853').setRanges([responseCountRange]).build(); //green
    var noneRespondedCF = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied("=G2>=C2").setBackground('#cc0100').setRanges([responseCountRange]).build(); //red
    var notEnoughRespondedCF = SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0).setBackground('#fbbc04').setRanges([responseCountRange]).build(); //yellow
    cfRules.push(...[emptyCF, allRespondedCF, noneRespondedCF, notEnoughRespondedCF]);
  }

  sheet.setConditionalFormatRules(cfRules);
}
