function myFunction() {
  var x = DriveApp.getFolderById('1TMgYVD_6PjYcU1Bl2wZ_xIbng2NetCdD');
  var files = x.getFiles();

  while (files.hasNext()){
    var file = files.next();
    Logger.log(Object.getOwnPropertyNames(file.getBlob()));
    
    
    // Logger.log(PdfViewer(file.getBlob()).getPageCount());

    // var doc = DocumentApp.getActiveDocument();
    // var blob = doc.getAs(‘application/pdf’);
    // var pdfViewer = PdfViewer.open(blob);
    // var pageCount = pdfViewer.getPageCount();
    // return pageCount;
  }
}
