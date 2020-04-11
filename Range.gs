//=============================================================================================
// Class Range
//

class Range {
  constructor(range, name = "", sheetName = "") {
    this.myRange = range;
    this.myName = name;
    this.mySheetName = sheetName;
    if (name !== "") name = name + " "; // Pad it to trace nicely
    if (sheetName !== "") sheetName = sheetName + "!"; // Pad it to trace nicely
    this.myTrace = `{Range ${sheetName}${name}${Range.trace(range)}}`;
    trace(`NEW ${this.myTrace}`);
  }
  
  clear() {
    trace(`clear ${this.trace}`);
    this.myRange.setValue("");
  }

  refresh() { // Reload the range - e.g. if it has changed
    trace(`${this.trace} refresh`);
    let newRange = Range.getByName(this.myName); 
    this.myRange = newRange.range;
    this.myTrace = newRange.trace;    
  }
  
  deleteExcessiveRows(rowsToKeep) {
    let height = this.height;
    trace(`${this.trace} deleteExcessiveRows rowsToKeep=${rowsToKeep} height=${this.height}`);
    if (height > rowsToKeep) {
      let startRow = this.row + rowsToKeep;
      let rowsToDelete = height - rowsToKeep;
      trace(`deleteRows startRow=${startRow} rowsToDelete=${rowsToDelete}`);
      this.sheet.deleteRows(startRow, rowsToDelete);
      this.refresh(); // Reload the range as it has changed now
    }
  }
  
  getNextRow() {
    let nextRow = null;
    return nextRow;
  }
  
  static getByName(rangeName, sheetName = "", spreadsheet = null) {
    spreadsheet = spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
    var sheet = (sheetName === "") ? null : spreadsheet.getSheetByName(sheetName);
    let range = spreadsheet.getRangeByName(rangeName);
    if (range === null && sheet !== null) {
      trace (`attempting to get named range from sheet ${sheetName}`);
      range = sheet.getRange(rangeName);
    }
    if (range === null) {
      Error.fatal(`Cannot find named range ${rangeName}`);
    }

    let newRange = new Range(range, rangeName, sheetName);
    trace(`Range.getByName ${rangeName} --> ${newRange.trace}`);
    return newRange;
  }

  static trace(range) {
    return `[${range.getSheet().getName()}!${range.getA1Notation()}]`;
  }

  get range()    { return this.myRange; }
  get trace()    { return this.myTrace; }
  get sheet()    { return this.myRange.getSheet(); }
  get values()   { return this.myRange.getValues(); }
  get height()   { return this.myRange.getHeight(); }
  get row()      { return this.myRange.getRow(); }
  get column()   { return this.myRange.getColumn(); }
}
