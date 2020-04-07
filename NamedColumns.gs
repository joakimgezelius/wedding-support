class NamedColumns {
  constructor(rowName, columnNamesRange) {
    if (typeof(columnNamesRange) === 'string') {
      columnNamesRange = CRange.getByName("EventDetailsColumnIds");
    }
    this.range = columnNamesRange;
    this.rowOffset = columnNamesRange.row;
    this.columnOffset = columnNamesRange.column;
    this.rowName = rowName;
    this.columnNumbers = {};
    let columnNames = columnNamesRange.values[0];
    let columnCount = columnNames.length;
    trace(`NamedColumns ${this.rowName} Row Offset: ${this.rowOffset} Column Offset: ${this.columnOffset} Columns: ${columnNames}`);
    for (var columnNumber = 0; columnNumber < columnCount; ++columnNumber) {
      let columnName = columnNames[columnNumber];
      if (columnName !== "") { // Only 
        this.columnNumbers[columnName] = columnNumber;
        trace(`Column ${columnNumber}: ${columnName}`);
      }
    }
  }
  
  getColumnNumber(columnName) {
    if (!columnName in this.columnNumbers) {
      Error.fatal(`Unknown ${this.rowName} column: ${columnName}`);
    }
    let columnNumber = this.columnNumbers[columnName];
    return columnNumber;
  }
  
  getAbsoluteColumnNumber(columnName) {
    return this.getColumnNumber(columnName) + this.columnOffset;
  }

  getColumnLetter() {
    return "A";
  }
}
