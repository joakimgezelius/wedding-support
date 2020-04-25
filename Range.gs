//=============================================================================================
// Class Range
//

class Range {
  constructor(range, name = "", sheetName = "") {
    this._range = range;
    this._name = name;
    this._sheetName = sheetName;
    this._sheet = this._range.getSheet();
    this._currentRowOffset = 0;
    if (name !== "") name = name + " "; // Pad it to trace nicely
    if (sheetName !== "") sheetName = sheetName + "!"; // Pad it to trace nicely
    this._trace = `{Range ${sheetName}${name}${Range.trace(range)}}`;
    trace(`NEW ${this._trace}`);
  }
  
  clear() {
    trace(`clear ${this.trace}`);
    this._range.setValue("");
  }

  static getByName(rangeName, sheetName = "", spreadsheet = null) {
    spreadsheet = spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
    let sheet = null;
    if (sheetName !== "") { // A sheet name is provided, if the named range cannot be found globally, a local named range will be attempted
      sheet = spreadsheet.getSheetByName(sheetName);
      trace (`getSheetByName sheet ${sheetName} -> ${sheet}`);
    }
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

  get range()            { return this._range; }
  get name()             { return this._name; }
  get trace()            { return this._trace; }
  get sheet()            { return this._sheet; }
  get values()           { return this._range.getValues(); }
  get height()           { return this._range.getHeight(); }
  get row()              { return this._range.getRow(); }    // Row number of the first row in the range
  get column()           { return this._range.getColumn(); } // Column number of the first column in the range
  get currentRow()       { return this._range.offset(this._currentRowOffset, 0, 1); } // A range of 1 row height
  get currentRowOffset() { return this._currentRowOffset; }

  // 
  //  Dynamic Range Features
  //
  
  refresh() { // Reload the range - e.g. if it has changed
    trace(`${this.trace} refresh`);
    let newRange = Range.getByName(this.name); 
    this._range = newRange.range;
    this._trace = newRange.trace;
  }
  
  deleteExcessiveRows(rowsToKeep) {
    trace(`${this.trace} deleteExcessiveRows rowsToKeep=${rowsToKeep} height=${this.height}`);
    if (this.height > rowsToKeep) {
      let startRow = this.row + rowsToKeep;
      let rowsToDelete = this.height - rowsToKeep;
      trace(`deleteRows startRow=${startRow} rowsToDelete=${rowsToDelete}`);
      this.sheet.deleteRows(startRow, rowsToDelete);
      this.refresh(); // Reload the range as it has changed now
    }
  }
  
  minimizeAndClear(callback = null) {
    this.deleteExcessiveRows(2); // Delete all but the first two rows
    this.clear();
    this._currentRowOffset = 0;
    if (callback !== null) this.format(callback);
  }
  
  rewind() {
    this._currentRowOffset = 0;
  }
  
  getNextRow() {
    let row = this.currentRow;
    ++this._currentRowOffset;
    return row;
  }

  getNextRowValues() {
    return this.getNextRow().getValues[0];
  }
    
  getNextRowAndExtend() {
    let row = this.currentRow;
    ++this._currentRowOffset;
    if (this.height - this._currentRowOffset < 2) { // Extend if we're 1 row from the end
      this._sheet.insertRowBefore(row.getRowIndex()+1);
      this.refresh();
    }
    return row;
  }
  
  getPreviousRow() {
    if (this._currentRowOffset > 0) { // Don't back up beyond beginning of range
      --this._currentRowOffset;
    }
    return this.currentRow;
  }

  trim() {
    let excess = this.height - this._currentRowOffset;
    let rowsToKeep = Math.max(this.height - excess, 2); // Never less than two lines left
    trace(`${this.trace} trim ${excess} lines (height: ${this.height} rowsToKeep: ${rowsToKeep})`);
    this.deleteExcessiveRows(rowsToKeep);
  }
  
  format(callback) {
    callback(this._range);
  }

  // Named Column Fetures
  
  loadColumnNames() {
    let columnNamesRange = this.range.offset(-1, 0, 1); // Get the one row above the range, we assume this row holds the column names
    this.namedColumns = new NamedColumns(this.name, columnNamesRange);
    return this;
  }

  getColumnOffset(columnName) {
    return this.namedColumns.getColumnOffset(columnName);
  }
  
  getColumnLetter(columnName) {
    return this.namedColumns.getColumnLetter(columnName);
  }
  
}


//----------------------------------------------------------------------------------------------------
// 
// class RangeRow

class RangeRow {

  constructor(data, rowOffset, containerRange) {
    this.data = data;
    this.rowOffset = rowOffset;
    this.containerRange = containerRange;
  }

  get(columnName, expectedType = "undefined") {
    let value = this.data[this.containerRange.getColumnOffset(columnName)];
    let actualType = typeof value;
    if (actualType !== expectedType) {
      switch (expectedType) {
        case "string": 
          return String(value); 
          break;
        case "date":
          break;
        case "number":
          value = Number(value);
          if (!Number.isNaN(value)) return value;
      }
      if (expectedType !== undefined && actualType !== expectedType) {
        let rowPosition = this.containerRange.row + this.rowOffset;
        let columnLetter = this.containerRange.getColumnLetter(columnName);
        Error.fatal(`Unexpected value in row ${rowPosition}, column ${columnLetter} (${columnName}), found a ${actualType} (${value}), expected a ${expectedType}`);
      }
    }
    return value;
  }

  getCell(columnName) {
    let columnOffset = this.containerRange.getColumnOffset(columnName);
    let cell = this.containerRange.range.offset(this.rowOffset, columnOffset, 1, 1);
//  trace(`${this.containerRange.name}.getCell ${columnName} --> ${Range.trace(cell)}`);
    return cell;
  }

  set(columnName, value) { 
    let cell = this.getCell(columnName);
    trace(`${this.containerRange.name}.set ${columnName} ${Range.trace(cell)} = ${value}`);
    cell.setValue(value);
  }

  getA1Notation(columnName) { 
    let cell = this.getCell(columnName);
    return cell.getA1Notation();
  }

}
