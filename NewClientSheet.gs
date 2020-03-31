function onCreateNewClient() {
  var name = Dialog.prompt('Create New Client Sheet', 'Please enter client sheet name:');
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


//=============================================================================================
// Class Sheet
//

var Sheet = function() {
}

Sheet.autofit = function(sheet) {
  trace("Sheet.autofit " + sheet.getName());
  sheet.autoResizeColumns(1, sheet.getMaxColumns());
  sheet.autoResizeRows(1, sheet.getMaxRows());
}

Sheet.getCleared = function(name) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(name);
  if (sheet == null) {
  }
  return sheet;
}

function saveRangeAsSpreadsheet(){ 
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var range = sheet.getRange('Sheet1!A1:B3');
  sheet.setNamedRange('buildingNameAddress', range);
  var TestRange = sheet.getRangeByName('buildingNameAddress').getValues(); 
  Logger.log(TestRange); 
  var destFolder = DriveApp.getFolderById("xxxxxxxxxxxxxxxxxxxxx"); 
  DriveApp.getFileById(sheet.getId()).makeCopy("Test File", destFolder); 
}

function saveAsSpreadsheet2() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var destFolder = DriveApp.getFolderById("xxxxxxxxxxxxxxxxx");
  DriveApp.getFileById(sheet.getId()).makeCopy("desired file name", destFolder);
} //END function saveAsSpreadsheet

function getParentFolder() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var file = DriveApp.getFileById(spreadsheet.getId());
  var folders = file.getParents();
  var folder = folders.next();  
  trace('getParentFolder --> ' + folder.getName());
  return folder;
}

function fileExists(name, folder) {
  var files = folder.getFilesByName(name);
  result = (files.hasNext()) ? true : false;
  trace("fileExists(" + folder.getName() + ", " + name + ") --> " + result);
  return result;
}
