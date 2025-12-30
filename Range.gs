//=============================================================================================
// Class Range
// Wrapper for https://developers.google.com/apps-script/reference/spreadsheet/range
//=============================================================================================

class Range {
  constructor(nativeRange, name = "") {
    this._nativeRange = nativeRange;
    this._name = name;
    this._sheet = new Sheet(nativeRange.getSheet());
    this._values = this.nativeRange.getValues();  // Note: we cache this as we need to allow the array to be sorted
    this._sheetName = this.sheet.name;
    this._currentRowOffset = 0;
    this._currentColumnOffset = 0;
    this._trace = `{Range ${name} ${Range.trace(nativeRange)} `; // NOTE: This static trace string is incomplete on purpose, see trace access method below
    trace(`NEW ${this.trace}`);
  }
  
  clear() {
    trace(`clear ${this.trace}`);
    this.nativeRange.setValue("");
    this._values = this.nativeRange.getValues();
  }

  // Get named range from active spreadsheet
  static getByName(rangeName, sheetName = "") {
    return Spreadsheet.active.getRangeByName(rangeName, sheetName);
  }

  static trace(nativeRange) {
    return `[${nativeRange.getSheet().getName()}!${nativeRange.getA1Notation()}]`;
  }

  getNativeRowRange(rowOffset) {
    return this.nativeRange.offset(rowOffset, 0, 1); // A range of 1 row height
  }

  extend(rows, columns = 0) {
    trace(`Range.extend ${this.trace} by ${rows} rows and ${columns} columns`);
    this._nativeRange = this.sheet.nativeSheet.getRange(this.rowPosition, this.columnPosition, this.height + rows, this.width + columns);
    this._values = this.nativeRange.getValues();
    return this;
  }

  copyTo(destination, copyPasteType) {
    trace(`Range.copyTo ${this.trace} to ${destination.trace}`);
    this.nativeRange.copyTo(destination.nativeRange, copyPasteType, false);
  }

  setName(name) {
    trace(`Range.setName ${this.trace} to ${name}`);
    let spreadsheet = Spreadsheet.active;
    spreadsheet.setNamedRange(name, this);
  }

  // Iterate over all rows in the range by moving the currentRowOffset forward
  //  - call back passing "this", callee picks up current row...
  //
  forEachRow(callback, context = null) {
    trace(`${this.trace} forEachRow`);
    const rowCount = this.height;
    let rowOffset = 0;
    for (rowOffset; rowOffset < rowCount; ++rowOffset) {
      this._currentRowOffset = rowOffset;
      let goOn = callback(this, context);
//    if (!goOn) break;
    }
  }

  get nativeRange()         { return this._nativeRange; }
  get name()                { return this._name; }
  get trace()               { return this._trace + this._currentRowOffset + "}"; }
  get sheet()               { return this._sheet; }
  get values()              { return this._values; }
  get value()               { return this._values[0][0]; } // Helper to access first cell in range
  get height()              { return this.nativeRange.getHeight(); }
  get width()               { return this.nativeRange.getWidth(); }
  get rowPosition()         { return this.nativeRange.getRow(); }    // Row number of the first row in the range
  get columnPosition()      { return this.nativeRange.getColumn(); } // Column number of the first column in the range
  get currentRow()          { return this.getNativeRowRange(this._currentRowOffset); } // A range of 1 row height
  get currentRowOffset()    { return this._currentRowOffset; }
  get currentRowValues()    { return this.values[this._currentRowOffset]; }
  get currentRowIsEmpty()   { return !this.currentRowValues.join(""); }
  get currentColumnOffset() { return this._currentColumnOffset; }

  set values(values)        { this.nativeRange.setValues(this._values = values); }
  set value(value)          { this.nativeRange.getCell(1, 1).setValue(value); }
  set currentRowOffset(value){ return this._currentRowOffset = value; }
  
  //========================================================================================================= 
  //  Dynamic Range Features Below
  //
  
  refresh() { // Reload the range - e.g. if it has changed
    trace(`${this.trace} refresh`);
    let newRange = this.sheet.getRangeByName(this.name); 
    this._nativeRange = newRange.nativeRange;
    this._values = this.nativeRange.getValues();
    this._trace = newRange.trace;
  }

  findSelectedRow() {
    let selection = this.sheet.activeRange;
    if (selection === null) { // Nothing is selected
      trace(`${this.trace} getSelectedRow, nothing selected`);
      return null;
    }
    let rowOffset = selection.rowPosition - this.rowPosition;
    if (rowOffset < 0 || rowOffset >= this.height) {  // Selected row is outside of range
      trace(`${this.trace} getSelectedRow, selection outside of range (row: ${selection.rowPosition})`);
      return null;
    }
    this._currentRowOffset = rowOffset;
    this._currentColumnOffset = selection.columnPosition - this.columnPosition;
    trace(`${this.trace} getSelectedRow --> row ${rowOffset}`);
    return this.currentRow;
  }

  findSelectedColumnOffset() {
    findSelectedRow();
    // TODO: ensure selection is within range
    return this.currentColumnOffset;
  }
  
  deleteExcessiveRows(rowsToKeep) {
    trace(`${this.trace} deleteExcessiveRows rowsToKeep=${rowsToKeep} height=${this.height}`);
    if (this.height > rowsToKeep) {
      const startRow = this.rowPosition + rowsToKeep;
      const rowsToDelete = this.height - rowsToKeep;
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
    return this._currentRowOffset;
  }

  findFirstEmptyRow() {
    this.rewind();
    return this.findNextEmptyRow();
  }
  
  findNextEmptyRow() {
    const values = this.values;
    while (values[this._currentRowOffset].join("")) { // While row is not empty
      ++this._currentRowOffset;
      if (this._currentRowOffset >= this.height) {
        trace(`${this.trace} findNextEmptyRow --> no more empty rows in range`);
        Error.fatal(`No more empty rows in range ${this.name}`);
        return null;
      }
    }
    trace(`${this.trace} findNextEmptyRow --> row ${this.currentRowOffset}`);
    return this.currentRow;
  }

  // Loop through all lines from the bottom until a non-empty row is found
  //
  findFirstTrailingEmptyRow() {
    const values = this.values;
    this._currentRowOffset = this.height - 1;
    while (this._currentRowOffset >= 0 && !values[this._currentRowOffset].join("")) {
      --this._currentRowOffset;
    }
    ++this._currentRowOffset;
    if (this._currentRowOffset == this.height) { // We're past the end, i.e. no empty lines left
      trace(`${this.trace} findFirstTrailingEmptyRow --> no more empty rows in range`);
      Error.fatal(`No more empty rows in range ${this.name}`);
      return null;
    }
    trace(`${this.trace} findFirstTrailingEmptyRow --> row ${this.currentRowOffset}`);
    return this.currentRow;
  }
  
  getNextRow() {
    const row = this.currentRow;
    ++this._currentRowOffset;
    return row;
  }

  getNextRowValues() {
    return this.getNextRow().getValues[0];
  }
    
  getNextRowAndExtend() {
    const row = this.currentRow;
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
    const excess = this.height - this._currentRowOffset;
    const rowsToKeep = Math.max(this.height - excess, 2); // Never less than two lines left
    trace(`${this.trace} trim ${excess} lines (height: ${this.height} rowsToKeep: ${rowsToKeep})`);
    this.deleteExcessiveRows(rowsToKeep);
  }
  
  format(callback) {
    callback(this.nativeRange);
  }

  // Named Column Features
  
  loadColumnNames() {
    if (this.namedColumns === undefined) {
      const columnNamesNativeRange = this.nativeRange.offset(-2, 0, 1); // Get the two row above the range, we assume this row holds the column names
      this.namedColumns = new NamedColumns(this.name, columnNamesNativeRange);
    }
    return this;
  }
  
  getColumnOffset(columnName) {
    return this.namedColumns.getColumnOffset(columnName);
  }
  
  getColumnLetter(columnName) {
    return this.namedColumns.getColumnLetter(columnName);
  }

} // Range

//=============================================================================================
// Class NamedRange
// Wrapper for https://developers.google.com/apps-script/reference/spreadsheet/named-range
//=============================================================================================

class NamedRange {
  constructor(nativeNamedRange) {
    this._nativeNamedRange = nativeNamedRange;
    this._range = null;
    try {
      this._range = nativeNamedRange.getRange();
    }
    catch(error) {
      trace(`Caught error ${error} while constructing NamedRange ${this.name}`);
    }
    const rangeString =  this._range === null ? '[#REF]' : `[${this._range.getA1Notation()}]`;
    this._trace = `{NamedRange ${this.name} ${rangeString}}`;
    trace(`NEW ${this.trace}`);
  }
  
  remove() {
    trace(`${this.trace}.remove`); 
    this._nativeNamedRange.remove();
  }

  get nativeNamedRange()    { return this._nativeNamedRange; }
  get name()                { return this._nativeNamedRange===null ? '[null]' : this._nativeNamedRange.getName(); }
  get range()               { return this._range; }
  get trace()               { return this._trace; }

  set name(name)            { return this._nativeNamedRange.setName(name) }
  set range(range)          { this._range = range; return this._nativeNamedRange.setRange(range) }

} // NamedRange


//=============================================================================================
// Class RangeRow
//=============================================================================================

class RangeRow {

  constructor(range, values = null) {
    this.values = (values === null) ? range.currentRowValues : values;
    this.namedColumns = range.namedColumns;
    this.nativeRange = range.currentRow;
    this.rowPosition = range.rowPosition + range.currentRowOffset;
//  trace(`NEW RangeRow on ${range.trace}`); // Avoid tracing unless critical for debugging, floods the trace.
  }

  get(columnName, expectedType = "undefined") {
    let value = this.values[this.namedColumns.getColumnOffset(columnName)];
    if (expectedType === "undefined") return value; // If we accept any type, just return the value!
    let actualType = typeof value;
    if (actualType !== expectedType) {
      switch (expectedType) {
        case "string": 
          return String(value); 
          break;
        case "date":
          break;
        case "boolean":
          return Boolean(value);
        case "number":
          value = Number(value);
          if (!Number.isNaN(value)) return value;
      }
      if (expectedType !== undefined && actualType !== expectedType) {
        let columnLetter = this.namedColumns.getColumnLetter(columnName);
        Error.fatal(`Unexpected value in row ${this.rowPosition}, column ${columnLetter} (${columnName}), found a ${actualType} (${value}), expected a ${expectedType}`);
      }
    }
    return value;
  }

  getCell(columnName) {
    const columnOffset = this.namedColumns.getColumnOffset(columnName);
    const cell = this.nativeRange.offset(0, columnOffset, 1, 1);
//  trace(`RangeRow.getCell ${columnName} --> ${Range.trace(cell)}`);
    return cell;
  }

  set(columnName, value) { 
    let cell = this.getCell(columnName);
    trace(`RangeRow.set ${columnName} ${Range.trace(cell)} = ${value}`);
    cell.setValue(value);
  }

/* Avoid this feature as it creates issues downstream, there is no simple way to get to the URL using standard sheet functions
  setHyperLink(columnName, url, label = url) {
    let value = `=hyperlink("${url}","${label}")`;
    trace(`RangeRow.setHyperLink ${columnName} = [${url}, ${label}]`);
    this.set(columnName, value);
  }
*/

  copyFieldTo(destination, columnName) {
    destination.set(columnName, this.get(columnName));
  }

  copyFieldsTo(destination, columnNames) {
    columnNames.forEach(columnName => this.copyFieldTo(destination, columnName)); 
  }

  getA1Notation(columnName) { 
    const cell = this.getCell(columnName);
    return cell.getA1Notation();
  }

} // RangeRow

