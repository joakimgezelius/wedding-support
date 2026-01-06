//----------------------------------------------------------------------------------------
// Wrapper for https://developers.google.com/apps-script/reference/spreadsheet/sheet
//
class Sheet {
  
  constructor(nativeSheet) {
    this._nativeSheet = nativeSheet;
    this._trace = `{Sheet "${this.name}"}`;
    trace(`NEW ${this.trace}`);
  }

  autofit() {
    trace("Sheet.autofit " + this.trace);
    this.nativeSheet.autoResizeColumns(1, this.maxColumns);
    this.nativeSheet.autoResizeRows(1, this.maxRows);
  }

  /*
  static saveRangeAsSpreadsheet() { 
    let sheet = SpreadsheetApp.getActiveSpreadsheet();
    let range = sheet.getRange("Sheet1!A1:B3");
    sheet.setNamedRange("buildingNameAddress", range);
    let TestRange = sheet.getRangeByName("buildingNameAddress").getValues(); 
    Logger.log(TestRange); 
    var destFolder = DriveApp.getFolderById("xxxxxxxxxxxxxxxxxxxxx"); 
    DriveApp.getFileById(sheet.getId()).copyTo(destFolder, "Test File"); 
  }
  */

  // Loop over all named ranges in the sheet and callback to the provided function
  iterateOverNamedRanges(callback=null) {
    // https://developers.google.com/apps-script/reference/spreadsheet/sheet#getnamedranges
    trace(`> ${this.trace}.iterateOverNamedRanges`);
    const namedRanges = this.nativeSheet.getNamedRanges();
    for (let i = 0; i < namedRanges.length; i++) {
      const namedRange = new NamedRange(namedRanges[i]);
      trace( `- iterateOverNamedRanges ${i}: ${namedRange.trace}`);
      if (callback !== null) callback(namedRange);
    }
    trace(`< ${this.trace}.iterateOverNamedRanges Done`);
  }

  // Get named range 
  // NOTE: returns a NamedRange object, not a Range - important difference!
  //
  getNamedRange(name) {
    const namedRanges = this.nativeSheet.getNamedRanges();
    for (let i = 0; i < namedRanges.length; i++) {
      const nativeNamedRange = namedRanges[i];
      if (nativeNamedRange.getName() === name) {
        const namedRange = new NamedRange(nativeNamedRange);
        trace(`${this.trace}.getNamedRange("${name}") -> ${namedRange.trace}`);
        return namedRange;
      }
    }
    trace(`${this.trace}.getNamedRange("${name}") not found -> null`);
    return null;
  }

  makeNamedRangesGlobal() {
    trace(`> ${this.trace}.makeNamedRangesGlobal`);
    this.iterateOverNamedRanges((namedRange) => {
      if (namedRange.name.includes('!')) { // We have a sheet-local named range
        trace(`Sheet.makeNamedRangesGlobal: Convert sheet-local named range ${namedRange.trace}`);
        namedRange.makeGlobal();
      }
    });
    trace(`< ${this.trace}.makeNamedRangesGlobal`);
  }

  copyTo(destination) {
    // https://developers.google.com/apps-script/reference/spreadsheet/sheet#copyTo(Spreadsheet)
    trace(`> ${this.trace}.copyTo("${destination.trace}`);
    const newNativeSheet = this.nativeSheet.copyTo(destination.nativeSpreadsheet);
    const newSheet = new Sheet(newNativeSheet);
    trace(`< ${this.trace}.copyTo("${destination.trace} --> ${newSheet.trace}`);
    return newSheet;
  }

  insertRowBefore(position)   { this.nativeSheet.insertRowBefore(position); return this; }
  deleteRows(position, count) { this.nativeSheet.deleteRows(position, count); return this; }
  activate()                  { this.nativeSheet.activate; return this; }
  
  get nativeSheet()      { return this._nativeSheet; }
  get name()             { return this.nativeSheet === null ? "NULL" : this.nativeSheet.getName(); }
  get maxColumns()       { return this.nativeSheet.getMaxColumns(); }
  get maxRows()          { return this.nativeSheet.getMaxRows(); }
  get selection()        { return this.nativeSheet.getSelection(); }
  get activeRange()      { return new Range(this.nativeSheet.getActiveRange()); } // The active range is the range selected
  get currentCell()      { return new Range(this.nativeSheet.getCurrentCell()); } // The current cell is the cell having focus
  get fullRange()        { return new Range(this.nativeSheet.getRange(1, 1, this.maxRows, this.maxColumns)); }
  get trace()            { return this._trace; }

  // https://developers.google.com/apps-script/reference/spreadsheet/sheet#setnamename
  set name(name)         { return this.nativeSheet.setName(name); }

} // Sheet


//----------------------------------------------------------------------------------------
// Wrapper for https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet
//

let _activeSpreadsheet = null;

class Spreadsheet {

  constructor(nativeSpreadsheet) {
    this._nativeSpreadsheet = nativeSpreadsheet;
    this._name = nativeSpreadsheet.getName();
    this._trace = `{Spreadsheet "${this._name}"}`;
    trace("NEW " + this.trace);
  }

  static get active() {
    return _activeSpreadsheet === null 
      ? _activeSpreadsheet = new Spreadsheet(SpreadsheetApp.getActiveSpreadsheet())
      : _activeSpreadsheet;
  } 

  static openById(id) {
    trace(`> Spreadsheet.openById(${id})`);
    // https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet-app#openById(String)
    let newSpreadsheet = SpreadsheetApp.openById(id);
    let spreadsheet = new Spreadsheet(newSpreadsheet);
    trace(`< Spreadsheet.openById(${id}) --> ${spreadsheet.trace}`);
    return spreadsheet;
  }

  static openByUrl(url) {
    trace(`> Spreadsheet.openByUrl(${url})`);
    // https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet-app#openByUrl(String)
    let newSpreadsheet = SpreadsheetApp.openByUrl(url);
    let spreadsheet = new Spreadsheet(newSpreadsheet);
    trace(`< Spreadsheet.openByUrl(${url}) --> ${spreadsheet.trace}`);
    return spreadsheet;
  }

  static get currentCellValue() {
    return SpreadsheetApp.getCurrentCell().getValues()[1,1];
  }

  static getCellValue(rangeName) {
    let nativeRange = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(rangeName);
    let value = nativeRange.getCell(1,1).getValue();
    trace(`Spreadsheet.getCellValue(${rangeName}) --> ${value}`);
    return value;
  }

  static getCellValueLinkUrl(rangeName) {
    let nativeRange = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(rangeName);
    // https://developers.google.com/apps-script/reference/spreadsheet/rich-text-value#getLinkUrl()
    let url = nativeRange.getCell(1,1).getRichTextValue().getLinkUrl();
    trace(`Spreadsheet.getCellValueLinkUrl(${rangeName}) --> ${url}`);
    return url;
  }

  setActive() {
    trace(`> Spreadsheet.setActiveSpreadsheet(${this.trace})`);
    // https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet-app#setActiveSpreadsheet(Spreadsheet)
    SpreadsheetApp.setActiveSpreadsheet(this.nativeSpreadsheet);
    trace(`< Spreadsheet.setActiveSpreadsheet`);
  }

  setActiveSheet(sheet) {
    // https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet#setactivesheetsheet
    trace(`${this.trace}.setActiveSheet(${sheet.trace})`);
    this.nativeSpreadsheet.setActiveSheet(sheet.nativeSheet);
    return sheet;
  }

  // Get range by name 
  // NOTE: returns a Range object, not a NamedRange - important difference! (see getNamedRange below)
  //
  getRangeByName(rangeName, failIfNotFound = true) {
    // https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet#getrangebynamename
    // 
    const nativeRange = this.nativeSpreadsheet.getRangeByName(rangeName);
    if (nativeRange !== null) { 
      // Range found, create wrapper and return
      const range = new Range(nativeRange, rangeName);
      trace(`${this.trace}.getRangeByName("${rangeName}") -> ${range.trace}`);
      return range;
    }
    if (failIfNotFound) {
      Error.fatal(`Cannot find range "${rangeName}"`);
    }
    trace(`${this.trace}.getRangeByName("${rangeName}") not found -> null`);
    return null;
  }

  // Get named range 
  // NOTE: returns a NamedRange object, not a Range - important difference! (see getRangeByName above)
  //
  getNamedRange(name) {
    const namedRanges = this.nativeSpreadsheet.getNamedRanges();
    for (let i = 0; i < namedRanges.length; i++) {
      const nativeNamedRange = namedRanges[i];
      if (nativeNamedRange.getName() === name) {
        const namedRange = new NamedRange(nativeNamedRange);
        trace(`getNamedRange("${name}") -> ${namedRange.trace}`);
        return namedRange;
      }
    }
    trace(`getNamedRange("${name}") not found -> null`);
    return null;
  }

  // Create a new named range
  //
  setNamedRange(name, range) {
    // https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet#setnamedrangename,-range
    trace(`${this.trace}.setNamedRange("${name}", ${range.trace})`);
    this.nativeSpreadsheet.setNamedRange(name, range.nativeRange);
  }

  removeNamedRange(name, forced) {
    // https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet#removenamedrangename
    trace(`${this.trace}.removeNamedRange(${name})`);
    if (forced || Dialog.confirm("Delete Named Range - Confirmation Required", `Are you sure you want to delete the named range ${name}?`) == true) {
      this.nativeSpreadsheet.removeNamedRange(name);
    }
  }

  // Loop over all named ranges in the spreadsheet and callback to the provided function
  //
  iterateOverNamedRanges(callback=null) {
    // https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet#getnamedranges
    trace(`> ${this.trace}.iterateOverNamedRanges`);
    const namedRanges = this.nativeSpreadsheet.getNamedRanges();
    for (let i = 0; i < namedRanges.length; i++) {
      const namedRange = new NamedRange(namedRanges[i]);
      trace( `- iterateOverNamedRanges ${i}: ${namedRange.trace}`);
      if (callback !== null) callback(namedRange);
    }
    trace(`< ${this.trace}.iterateOverNamedRanges Done`);
  }

  getSheetByName(name) {
    const sheet = this.nativeSpreadsheet.getSheetByName(name);
    const newSheet = sheet === null ? null : new Sheet(sheet);
    trace(`${this.trace}.getSheetByName("${name}") --> ${sheet === null?"null (NOT FOUND)":newSheet.trace}`);
    return newSheet;
  }
  
  getSheetPosition(sheet) {
    // https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet#getsheets
    trace(`${this.trace}.getSheetPosition(${sheet.trace})`);
    const sheets = this.nativeSpreadsheet.getSheets();
    const sheetName = sheet.name;
    for (let i = 0; i < sheets.length; i++) {
      if (sheets[i].getName() == sheetName) {
        trace(`${this.trace}.getSheetPosition(${sheet.trace}) --> ${i}`);
        return i;
      };
    }
    trace(`${this.trace}.getSheetPosition(${sheet.trace}) --> -1, Not Found`);
    return -1;
  }

  deleteSheet(sheet, forced) {
    // https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet#deletesheetsheet
    trace(`> ${this.trace}.deleteSheet(${sheet.trace}), forced=${forced}`);
    if (forced || Dialog.confirm("Delete Sheet Tab - Confirmation Required", `Are you sure you want to delete the sheet tab "${sheet.name}"?`) == true) {
      this.nativeSpreadsheet.deleteSheet(sheet.nativeSheet);
      trace(`- ${this.trace}.deleteSheet(${sheet.trace}) Done`);
    }
  }

  copy(name) {
    trace(`${this.trace}.copy(${name})`);
    const newSpreadsheet = new Spreadsheet(this.nativeSpreadsheet.copy(name));
    return newSpreadsheet;
  }
  
  get nativeSpreadsheet() { return this._nativeSpreadsheet; }
  get activeSheet()       { return new Sheet(this.nativeSpreadsheet.getActiveSheet()); }
  get name()              { return this._name; }
  get id()                { return this.nativeSpreadsheet.getId(); }
  get file()              { return new File(DriveApp.getFileById(this.id)); }
  get url()               { return this.nativeSpreadsheet.getUrl(); }
  get parentFolder()      { return this.file.parent; }
  get trace()             { return this._trace; }



} // Spreadsheet
