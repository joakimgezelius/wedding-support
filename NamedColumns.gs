class NamedColumns {

  constructor(rowName, columnNamesNativeRange) {
    this.baseRowPosition = columnNamesNativeRange.getRow();
    this.baseColumnPosition = columnNamesNativeRange.getColumn();
    this.rowName = rowName;
    this.columnOffsets = {};
    this.columnLetters = {};
    let columnNames = columnNamesNativeRange.getValues()[0];
    let columnCount = columnNames.length;
    trace(`NamedColumns ${this.rowName} Base Row Position: ${this.baseRowPosition} Base Column Position: ${this.baseColumnPosition} Columns: ${columnNames}`);
    for (var columnOffset = 0; columnOffset < columnCount; ++columnOffset) {
      let columnName = columnNames[columnOffset];
      if (columnName !== "") { // Only 
        this.columnOffsets[columnName] = columnOffset;
        let baseColumnOffset = this.baseColumnPosition - 1; // Zero-based
        let globalColumnOffset = baseColumnOffset + columnOffset
        let letter = globalColumnOffset > 25 /*>Z?*/ ? "A" + String.fromCharCode(65 - 26 + globalColumnOffset) : String.fromCharCode(65 /* ascii('A') */ + globalColumnOffset);
        this.columnLetters[columnName] = letter;
        //trace(`Column ${columnOffset}: ${columnName}`);
      }
    }
    //console.log(`NamedColumns ${this.rowName} Column offsets: \n`, this.columnOffsets);
    //console.log(`NamedColumns ${this.rowName} Column letters: \n`, this.columnLetters);
  }
  
  getColumnOffset(columnName) {
    if (!(columnName in this.columnOffsets)) {
      Error.fatal(`Unknown ${this.rowName} column: ${columnName}`);
    }
    let offset = this.columnOffsets[columnName];
    //trace(`getColumnOffset(${columnName}) --> ${offset}`);
    return offset;
  }
  
  getAbsoluteColumnOffset(columnName) {
    return this.getColumnOffset(columnName) + this.baseColumnPosition - 1;
  }

  getAbsoluteColumnPosition(columnName) {
    return this.getColumnOffset(columnName) + this.baseColumnPosition;
  }

  getColumnLetter(columnName) {
    if (!(columnName in this.columnLetters)) {
      Error.fatal(`Unknown ${this.rowName} column: ${columnName}`);
    }
    let letter = this.columnLetters[columnName];
    //trace(`getColumnLetter(${columnName}) --> ${letter}`);
    return letter;
  }

}
