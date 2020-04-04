//=============================================================================================
// Class CellRange
//

class CellRange {
  constructor(range) {
    this.myRange = range;
  }
  
  clear() {
    trace(`clear ${this.trace}`);
    Range.clear(this.range);
  }

  static getByName(rangeName, spreadsheet) {
    let cellRange = new CellRange(Range.getByName(rangeName, spreadsheet));
    trace(`CellRange.getByName ${rangeName} --> ${cellRange.trace}`);
    return cellRange;
  }

  get range()  { return this.myRange; }
  get trace()  { return `{CellRange ${Range.trace(this.range)}}`; }
  get sheet()  { return this.myRange.getSheet(); }
  get values() { return this.myRange.getValues(); }
  get height() { return this.myRange.getHeight(); }
}


//=============================================================================================
// Class Range
//

class Range {
  static clear(range) {
    trace(`Range.clear  ${Range.trace(range)}`);
    range.setValue("");
  }

  static getByName(rangeName, spreadsheet) {
    if (spreadsheet == null) {
      spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    }
    let range = spreadsheet.getRangeByName(rangeName)
    if (range === null) {
      Error.fatal(`Cannot find named range ${rangeName}`);
    }
    trace(`Range.getByName ${rangeName} --> ${Range.trace(range)}`);
    return range;
  }

  static trace(range) {
    return `[${range.getSheet().getName()}!${range.getA1Notation()}]`;
  }
}
