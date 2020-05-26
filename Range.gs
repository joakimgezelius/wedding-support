//=============================================================================================
// Class Range
//

class Range {
  constructor(nativeRange, name = "") {
    this._nativeRange = nativeRange;
    this._name = name;
    this._sheet = new Sheet(nativeRange.getSheet());
    this._sheetName = this.sheet.name;
    this._currentRowOffset = 0;
    this._trace = `{Range ${name} ${Range.trace(nativeRange)}}`;
    trace(`NEW ${this._trace}`);
  }
  
  clear() {
    trace(`clear ${this.trace}`);
    this.nativeRange.setValue("");
  }

  static getByName(rangeName, sheetName = "") {
    return Spreadsheet.active.getRangeByName(rangeName, sheetName);
  }
  
  static trace(range) {
    return `[${range.getSheet().getName()}!${range.getA1Notation()}]`;
  }

  get nativeRange()      { return this._nativeRange; }
  get name()             { return this._name; }
  get trace()            { return this._trace; }
  get sheet()            { return this._sheet; }
  get values()           { return this.nativeRange.getValues(); }
  get formulas()         { return this.nativeRange.getFormulas(); }
  get height()           { return this.nativeRange.getHeight(); }
  get rowPosition()      { return this.nativeRange.getRow(); }    // Row number of the first row in the range
  get columnPosition()   { return this.nativeRange.getColumn(); } // Column number of the first column in the range
  get currentRow()       { return this.nativeRange.offset(this.currentRowOffset, 0, 1); } // A range of 1 row height
  get currentRowOffset() { return this._currentRowOffset; }

  // 
  //  Dynamic Range Features
  //
  
  refresh() { // Reload the range - e.g. if it has changed
    trace(`${this.trace} refresh`);
    let newRange = this.sheet.getRangeByName(this.name); 
    this._nativeRange = newRange.nativeRange;
    this._trace = newRange.trace;
  }
  
  deleteExcessiveRows(rowsToKeep) {
    trace(`${this.trace} deleteExcessiveRows rowsToKeep=${rowsToKeep} height=${this.height}`);
    if (this.height > rowsToKeep) {
      let startRow = this.rowPosition + rowsToKeep;
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
      this.sheet.insertRowBefore(row.getRowIndex()+1);
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
    callback(this.nativeRange);
  }

  // Named Column Features
  
  loadColumnNames() {
    if (this.namedColumns === undefined) {
      let columnNamesRange = this.nativeRange.offset(-1, 0, 1); // Get the one row above the range, we assume this row holds the column names
      this.namedColumns = new NamedColumns(this.name, columnNamesRange);
    }
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

  constructor(values, formulas, rowOffset, containerRange) {
    this.values = values;
    this.formulas = formulas;
    this.rowOffset = rowOffset;
    this.containerRange = containerRange;
  }

  get(columnName, expectedType = "undefined") {
    let value = this.values[this.containerRange.getColumnOffset(columnName)];
    if (expectedType === "undefined") return value; // If we accept any type, just return the value!
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
        let columnLetter = this.containerRange.getColumnLetter(columnName);
        Error.fatal(`Unexpected value in row ${this.rowPosition}, column ${columnLetter} (${columnName}), found a ${actualType} (${value}), expected a ${expectedType}`);
      }
    }
    return value;
  }

  getFormula(columnName) {
    let formula = this.formulas[this.containerRange.getColumnOffset(columnName)];
    return formula;
  }
  
  getCell(columnName) {
    let columnOffset = this.containerRange.getColumnOffset(columnName);
    let cell = this.containerRange.nativeRange.offset(this.rowOffset, columnOffset, 1, 1);
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

  get rowPosition() {
    return this.containerRange.rowPosition + this.rowOffset;
  }

}
