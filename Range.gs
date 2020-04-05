//=============================================================================================
// Class CRange
//

class CRange {
  constructor(range, name) {
    this.myRange = range;
    this.myName = name;
    if (name !== "") name = name + " "; 
    this.myTrace = `{CRange ${name}${CRange.trace(range)}}`;
    trace(`NEW ${this.myTrace}`);
  }
  
  clear() {
    trace(`clear ${this.trace}`);
    this.myRange.setValue("");
  }

  refresh() { // Reload the range - e.g. if it has changed
    trace(`${this.trace} refresh`);
    let newRange = CRange.getByName(this.myName); 
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
  
  static getByName(rangeName, spreadsheet) {
    if (spreadsheet == null) {
      spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    }
    let range = spreadsheet.getRangeByName(rangeName)
    if (range === null) {
      Error.fatal(`Cannot find named range ${rangeName}`);
    }
    let newRange = new CRange(range, rangeName);
    trace(`CRange.getByName ${rangeName} --> ${newRange.trace}`);
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
}
