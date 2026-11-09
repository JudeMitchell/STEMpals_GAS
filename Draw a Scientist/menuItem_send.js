function myFunction() {
  var x = DriveApp.getFiles();
  while ( x.hasNext() ){
    var y = x.next();
    
    if ( /png/.test(y.getName())){
      Logger.log(y.getName());
      Logger.log(y.getId());
    }
    
  }
}
