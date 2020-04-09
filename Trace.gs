// Global/static trace variables
//
var useTrace = true;
var traceRow = 1;
var currentSpreadSheet = SpreadsheetApp.getActiveSpreadsheet();
var traceArea = null;

function trace(text) {
  if (useTrace == true) {
    console.log(text);
  /*  
    if (traceArea == null) { // Not initialised yet
      var sheet = currentSpreadSheet.getSheetByName("Trace");
      if (sheet == null) { // Add "Trace" sheet at the end if it doesn't exist
        sheet = currentSpreadSheet.insertSheet("Trace", 99);
      }
      traceArea = sheet.getRange(1, 1, sheet.getMaxRows());
      Trace.clear(); // Always clear trace upon new run
    }
    traceArea.getCell(traceRow++,1).setValue(text);
    */
  }
}

class Trace {
  
  static clear() {
    // Clear trace area
    traceArea.setValue("");
    trace("Trace cleared: " + CRange.trace(traceArea));
  }

/*
  static showTraceSidebar() {
    var html = HtmlService.createHtmlOutputFromFile('Page')
      .setTitle('My custom sidebar')
      .setWidth(300);
    SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .showSidebar(html);
  }
*/
}
