//=============================================================================================
// Class Error

class Error {
  constructor(message) {
    this._message = message;
  }

  throw() {
    throw(this._message);    
  }
  
  static fatal(errorMessage) {
    let error = new Error(errorMessage);
    trace("Fatal error, terminating: " + errorMessage);
    //  Browser.msgBox("Fatal error:" + errorMessage, Browser.Buttons.OK);
    error.throw();
  }

  get message() { return this._message; }
  
  static get break() {
    trace("BREAK");
    throw("Break");
  }
}
