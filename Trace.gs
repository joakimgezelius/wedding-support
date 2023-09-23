// Global/static trace variables
//
const useTrace = true;

function trace(text) {
  if (useTrace == true) {
    console.log(text);
  }
}

class Trace {
  
  static clear() {
    // Clear trace area
  }

  static traceObject(object) {                          // Unpacks an object to string and gives type of an object
    let results = JSON.stringify(object);
    let type = typeof(results);
    trace(`results = ${results} ---> object type = ${type}`);    
  }

/*
  static showTraceSidebar() {
    var html = HtmlService.createHtmlOutputFromFile("Page")
      .setTitle("My custom sidebar")
      .setWidth(300);
    SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .showSidebar(html);
  }
*/
}

class Log {
  static clear() {

  }  

  static log() {

  }
}
