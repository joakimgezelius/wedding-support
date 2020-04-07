//=============================================================================================
// Class Error

class Error {
  constructor(message) {
    this.myMessage = message;
  }

  throw() {
    throw(this.myMessage);    
  }
  
  static fatal(errorMessage) {
    let error = new Error(errorMessage);
    trace("Fatal error, terminating: " + errorMessage);
    //  Browser.msgBox("Fatal error:" + errorMessage, Browser.Buttons.OK);
    error.throw();
  }

  get message() { return this.myMessage; }
  
  static get break() {
    trace("BREAK");
    throw("Break");
  }
}
