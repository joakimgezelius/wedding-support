class Sheet {
  
  constructor(sheet) {
    this._sheet = sheet;
    this._trace = `{Sheet "${this.name}"}`;
    trace(`NEW ${this.trace}`);
  }

  autofit() {
    trace("Sheet.autofit " + this.trace);
    sheet.autoResizeColumns(1, this.maxColumns);
    sheet.autoResizeRows(1, this.maxRows);
  }

  /*
  static getCleared(name) {
    let sheet = Spreadsheet.active.getSheetByName(name);
    if (sheet === null) {
    }
    return sheet;
  }
  */
  /*
  static saveRangeAsSpreadsheet() { 
    let sheet = SpreadsheetApp.getActiveSpreadsheet();
    let range = sheet.getRange("Sheet1!A1:B3");
    sheet.setNamedRange("buildingNameAddress", range);
    let TestRange = sheet.getRangeByName("buildingNameAddress").getValues(); 
    Logger.log(TestRange); 
    var destFolder = DriveApp.getFolderById("xxxxxxxxxxxxxxxxxxxxx"); 
    DriveApp.getFileById(sheet.getId()).makeCopy("Test File", destFolder); 
  }
  */

  getRangeByName(name) {
    range = this.sheet.getRange(rangeName);
    newRange = range === null ? null : new Range(range, name, this.name);
    trace(`${this.trace}.getRangeByName("${name}") --> ${range === null?"null (NOT FOUND)":newSheet.trace}`);
    return newRange;
  }

  insertRowBefore(position)   { this._sheet.insertRowBefore(position); return this; }
  deleteRows(position, count) { this._sheet.deleteRows(position, count); return this; }
  copyTo(destination)         { this._sheet.copyTo(destination); return this; }
  activate()                  { this._sheet.activate; return this; }
  
  get sheet()       { return this._sheet; }
  get name()        { return this._sheet === null ? "NULL" : this._sheet.getName(); }
  get maxColumns()  { return this._sheet.getMaxColumns(); }
  get maxRows()     { return this._sheet.getMaxRows(); }
  get selection()   { return this._sheet.getSelection(); }
  get activeRange() { return new Range(this._sheet.getActiveRange()); }
  get trace()       { return this._trace; }

} // Sheet


class Spreadsheet {

  constructor(spreadsheet) {
    this._spreadsheet = spreadsheet;
    this._name = spreadsheet.getName();
    this._trace = `{Spreadsheet "${this._name}"}`;
    trace("NEW " + this.trace);
  }

  static get active() {
    return new Spreadsheet(SpreadsheetApp.getActiveSpreadsheet());
  }

  static openById(id) {
    trace(`Spreadsheet.openById(${id})`);
    let newSpreadsheet = SpreadsheetApp.openById(id);
    return new Spreadsheet(newSpreadsheet);
  }
  
  getRangeByName(rangeName, sheetName = "") {
    let range = this.spreadsheet.getRangeByName(rangeName);
    if (range !== null) { 
      // Range found, create wrapper and return
      let newRange = new Range(range, rangeName, sheetName);
      trace(`${this.trace}.getRangeByName("${rangeName}") --> ${newRange.trace}`);
      return newRange;
    }
    if (sheetName !== "") {
      // A sheet name is provided and the named range cannot be found globally, a local named range will be attempted
      trace (`attempting to get named range from sheet "${sheetName}"`);
      let sheet = this.getSheetByName(sheetName);
      if (sheet !== null) {
        range = sheet.getRangeByName(rangeName);
        if (range !== null) {
          return range;
        }
      }
    }
    Error.fatal(`Cannot find named range ${rangeName}`);
  }

  getSheetByName(name) {
    let sheet = this.spreadsheet.getSheetByName(name);
    let newSheet = sheet === null ? null : new Sheet(sheet);
    trace(`${this.trace}.getSheetByName("${name}") --> ${sheet === null?"null (NOT FOUND)":newSheet.trace}`);
    return newSheet;
  }
  
  copy(name) {
    trace(`${this.trace}.copy(${name})`);
    let newSpreadsheet = new Spreadsheet(this.spreadsheet.copy(name));
    return newSpreadsheet;
  }
  
  get spreadsheet()      { return this._spreadsheet; }
  get activeSheet()      { return new Sheet(this.spreadsheet.getActiveSheet()); }
  get name()             { return this._name; }
  get id()               { return this.spreadsheet.getId(); }
  get file()             { return new File(DriveApp.getFileById(this.spreadsheet.id)); }
  get parentFolder()     { return this.file.parent; }
  get trace()            { return this._trace; }

} // Spreadsheet

