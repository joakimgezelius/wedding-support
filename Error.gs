//=============================================================================================
// Class Error

var Error = function() {
}

Error.fatal = function(errorMessage) {
  trace("Fatal error, terminating: " + errorMessage);
  //  Browser.msgBox("Fatal error:" + errorMessage, Browser.Buttons.OK);
  throw(errorMessage);
}
