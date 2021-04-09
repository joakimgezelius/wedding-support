WeddingClientTemplateSpreadsheetId = "1IAAbpD9bwZThh78ohFIJDAr9ebq9IM9obA28h3lGRDA";


function onCreateNewClientSheet() {
  var name = Dialog.prompt("Create New Client Sheet", "Please enter client sheet name:");
  if (name !== "CANCEL") {
    if (name === "") {
      throw("New client name required");
    }
    folder = getParentFolder();
    
    trace("Create duplicate spreadsheet '" + name + "' in folder '" + folder.getName() + "'");
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (fileExists(name, folder)) {
      throw("Spreadsheet already exists!");
    }
    DriveApp.getFileById(spreadsheet.getId()).makeCopy(name, folder);
  }
}

function onClearSheet() {
  trace("onClearSheet");
}



function saveAsSpreadsheet2() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var destFolder = DriveApp.getFolderById("xxxxxxxxxxxxxxxxx");
  DriveApp.getFileById(sheet.getId()).makeCopy("desired file name", destFolder);
} //END function saveAsSpreadsheet


