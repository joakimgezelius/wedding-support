function onCreateNewClient() {
  var name = Dialog.prompt("Create New Client Sheet", "Please enter client sheet name:");
  if (name !== "CANCEL") {
    if (name === "") {
      throw("New client name required");
    }
    folder = getParentFolder();
    
    trace("Create duplicate spreadsheet '" + name + "' in folder '" + folder.getName() + "'");
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (fileExists(name, folder)) {
      throw("Spreadsheet already exists");
    }
    DriveApp.getFileById(spreadsheet.getId()).makeCopy(name, folder);
  }
}

function onClearSheet() {
  trace("onClearSheet");
}

function getParentFolder() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var file = DriveApp.getFileById(spreadsheet.getId());
  var folders = file.getParents();
  var folder = folders.next();  
  trace("getParentFolder --> " + folder.getName());
  return folder;
}

function saveAsSpreadsheet2() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var destFolder = DriveApp.getFolderById("xxxxxxxxxxxxxxxxx");
  DriveApp.getFileById(sheet.getId()).makeCopy("desired file name", destFolder);
} //END function saveAsSpreadsheet

function fileExists(name, folder) {
  var files = folder.getFilesByName(name);
  result = (files.hasNext()) ? true : false;
  trace("fileExists(" + folder.getName() + ", " + name + ") --> " + result);
  return result;
}


//=============================================================================================
// Class Sheet
//

class Sheet {
  
  constructor(sheet) {
    mySheet = sheet;
    myTrace = `{Sheet ${this.name}}`;
    trace(`NEW ${this.trace}`);
  }

  autofit() {
    trace("Sheet.autofit " + this.name);
    sheet.autoResizeColumns(1, this.maxColumns);
    sheet.autoResizeRows(1, this.maxRows);
  }

  static getCleared(name) {
    let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = spreadsheet.getSheetByName(name);
    if (sheet == null) {
    }
    return sheet;
  }

  static saveRangeAsSpreadsheet(){ 
    let sheet = SpreadsheetApp.getActiveSpreadsheet();
    let range = sheet.getRange("Sheet1!A1:B3");
    sheet.setNamedRange("buildingNameAddress", range);
    let TestRange = sheet.getRangeByName("buildingNameAddress").getValues(); 
    Logger.log(TestRange); 
    var destFolder = DriveApp.getFolderById("xxxxxxxxxxxxxxxxxxxxx"); 
    DriveApp.getFileById(sheet.getId()).makeCopy("Test File", destFolder); 
  }
  
  get sheet()      { return this.mySheet; }
  get name()       { return this.mySheet.GetName(); }
  get maxColumns() { return this.mySheet.MaxColumnsgetMaxColumns(); }
  get maxRows()    { return this.mySheet.getMaxRows(); }
  get trace()      { return this.myTrace; }
}
