function setRichTextCell(id, cellRange, type, text=''){
  if ( type == 'case' ){
    var url = "https://thenappycollective.lightning.force.com/lightning/r/Case/" + id + "/view";
    if ( text == '' ){
      var text = id;
    }
  } else if ( type == 'url' ){
    var url = id;
  } else if ( type == 'spreadsheet' ){
    var url = "https://docs.google.com/spreadsheets/d/" + id;
    if ( text == '' ){
      var text = id;
    }
  } else if ( type == 'folder' ){
    var url = "https://drive.google.com/drive/folders/" + id;
    if ( text == '' ){
      var text = id;
    }
  } else {
    Logger.log('setRichTextCell does not have url structure for type: ' + type);
    jude_fail();
  }

  var richText = SpreadsheetApp.newRichTextValue().setText(text).setLinkUrl(url).build();
  cellRange.setRichTextValue(richText);
}


function getIdFromUrl(url){
  if ( /spreadsheets/.test(url) ){
    var frontTrimmedUrl = url.replace('https://docs.google.com/spreadsheets/d/','');
    var slashIndex = frontTrimmedUrl.indexOf('/');
    if ( slashIndex != -1 ){
      return frontTrimmedUrl.slice(0,slashIndex);
    } else {
      return frontTrimmedUrl;
    }

  } else if (/folders/.test(url)) {
    return url.replace('https://drive.google.com/drive/folders/', '');
  } else {
    Logger.log('Cannot extract ID from URL: ' + url);
    jude_fail();
  }
}