const EnquiriesRangeName = "Enquiries";
const ParamsSheetName = "Params";
const CoordinationSheetName = "Coordination";

// Go through raw list of enquiries, 
//  - find those that are work-in-progress
//
function onUpdateEnquiries() {
  trace("onUpdateEnquiries");
  //let enquiries = new Enquiries;  
  //enquiries.update(enquiriesNoReply);
}

// Open client sheet for the selected client, 
// - create a new sheet if no such sheet exists
//
function onOpenClientSheet() {
  trace("onOpenClientSheet");
  let enquiries = new Enquiries;
  enquiries.selected.openClientSheet();
}

function onCreateNewClientSheet() {
  trace("onCreateNewClientSheet");
  let enquiries = new Enquiries;
  enquiries.selected.createNewClientSheet();
}

function onDraftSelectedEmail() {
  trace("onDraftSelectedEmail");
  let enquiries = new Enquiries;
  enquiries.selected.draftSelectedEmail();
}

// To Prepare Client Document Structure for Small Weddings

function onPrepareClientStructureSmallWedding() {
  trace("onPrepareClientStructureSmallWedding");
  let enquiries = new Enquiries;        
  let sourceFolderId = "1lIUlRJFAxoVsOy_Tdmga9ZqzTWZGDDxr";                                              // Source - Client Template Small Weds
  let templateClientSheetLinkCell = Range.getByName("TemplateClientSheetSmallWedding", ParamsSheetName); // Small wedding template reference
  enquiries.selected.prepareClientStructure(sourceFolderId,templateClientSheetLinkCell);  
}

// To Prepare Client Document Structure for Large Weddings

function onPrepareClientStructureLargeWedding() {
  trace("onPrepareClientStructureLargeWedding");
  let enquiries = new Enquiries;
  let sourceFolderId = "1EDg-xJDwsj-h68L_yFzuo5rVrHQHEB1t";                                              // Source - Client Template Large Weds
  let templateClientSheetLinkCell = Range.getByName("TemplateClientSheetLargeWedding", ParamsSheetName); // Large weddingtemplate reference
  enquiries.selected.prepareClientStructure(sourceFolderId,templateClientSheetLinkCell);
}

//========================================================================================================

class Enquiries {

  constructor(rangeName = null) {
    rangeName = (rangeName === null) ? EnquiriesRangeName : rangeName;
    this.range = Range.getByName(rangeName).loadColumnNames();
    trace("NEW " + this.trace);
  }

  get selected() {
    let selection = this.range.findSelectedRow();
    if (selection === null) {
      Error.fatal("Please select a valid enquiry.");
    }
    // Range is now positioned to selection
    let selectedEnquiry = new Enquiry(this.range);
    if (!selectedEnquiry.isValid) {
      Error.fatal("Please select a valid enquiry.");
    }
    return selectedEnquiry;
  }
  
  update(target) {
    trace(`${this.trace}.update`);
    for (var rowOffset = 0; rowOffset < this.range.height; rowOffset++) {
      let enquiry = new Enquiry(); // Fix this!!
      if (!enquiry.isValid) {
        break;
      }
      target.append(enquiry);
    }
  }
  
  // Append an enquiry to a list 
  //
  append(enquiry) {
    trace(`${this.trace}.append ${enquiry.trace}`);    
    this.range.findFirstTrailingEmptyRow();
//  trace(`Create target enquiry based on ${this.range.trace}, ${this.range.currentRowOffset}`);    
    let targetEnquiry = new Enquiry(); // Fix this!!
    enquiry.copyTo(targetEnquiry);
  }

  get trace() { return `{Enquiries ${this.range.trace}}`; }
  
} // Enquiries


//================================================================================================


class Enquiry extends RangeRow {
  
  constructor(range) {
    super(range);
    this._name = this.name;
    this._rowOffset = range.currentRowOffset;
    this._isValid = (this.name !== "");
    trace("NEW " + this.trace);
  }
  
  createNewClientSheet() {
    trace(`createNewClientSheet for ${this.trace}`);
    if (!this.isValid) {
      Error.fatal("Please select a valid enquiry.");
    };
    let clientSheetName = `${this.name} (Prospect Client)`;
//  let targetFolder = Folder.getById(enquiriesFolderId);
    let weddingClientTemplateFile = File.getById(weddingClientTemplateSpreadsheetId);
    let clientSpreadsheetFile = weddingClientTemplateFile.copyTo(targetFolder, clientSheetName);
    this.clientSheet = Spreadsheet.openById(clientSpreadsheetFile.id);
    this.sheetId = this.clientSheet.id;
    this.sheetLink = `=hyperlink("${this.clientSheet.url}";"...")`;
    Browser.newTab(this.clientSheet.url);
  }

  prepareClientStructure(sourceFolderId,templateClientSheetLinkCell) {
    trace(`prepareClientStructure for ${this.trace}`);
    let destinationFolderId = "19y3-Zou_RAWHZKaZ_5W_FJXql_Pz-gdd";   // Destination - W & E's
    let sourceFolder = Folder.getById(sourceFolderId);
    let destinationFolder = Folder.getById(destinationFolderId);    
    if (destinationFolder.folderExists(this.fileName) == true) {
      Dialog.notify("Client folder already exists","Please check the Weddings & Events Folder for more details.");      
    } else {
      Dialog.notify("Preparing the Structure", "Making the new Client Document Structure, This may take a few seconds...");
      sourceFolder.copyTo(destinationFolder, this.fileName);                 // Copies source folder contents to destination folder with given name

      let templateClientSheetLink = templateClientSheetLinkCell.nativeRange.getRichTextValue().getLinkUrl();
      let templateSheetFile = File.getByUrl(templateClientSheetLink);
      trace(` Client Template Sheet: ${templateSheetFile.name}`);

      let clientFolder = destinationFolder.getFolderByName(this.fileName);    // Gets newly created client folder by name
      trace(` New Client Folder: ${clientFolder}`);
      let clientFolderLink = clientFolder.url;                                // Gets the URL of newly created client folder  
      trace(` New Client Folder Link: ${clientFolderLink}`);   
      this.set("FolderLink",clientFolderLink);                                // Places the client folder link to the FolderLink Column
    
      /*let targetFolderName = "Office Use";                                  // Folder name to copy the template file in it
      let subFolder = clientFolder.getFolderByName(targetFolderName);
      trace(` Sub Folder: ${subFolder}`);
      let subFolderId = subFolder.getId();                                    // Gets the id of children folder with given folder name
    
      let targetFolder = Folder.getById(subFolderId);                         // Gets the folder by id to copy the template file in it
      if (targetFolder) {
         templateSheetFile.copyTo(targetFolder,this.fileName);       
         let newClientSheet = targetFolder.getFileByName(this.fileName);      // Gets the file with given name
         let newClientSheetId = newClientSheet.getId();                       // Returns the id of found file
         let clientTemplate = Spreadsheet.openById(newClientSheetId);
         clientTemplate.setActive();                                          // Sets the active spreadsheet Confirmed W & E to new client sheet
         Range.getByName("EventDetailsSectionA",CoordinationSheetName).clear();
         Range.getByName("EventDetailsSectionB",CoordinationSheetName).clear();
         Range.getByName("EventDetailsSectionC",CoordinationSheetName).clear();
         Range.getByName("EventDetailsSectionD",CoordinationSheetName).clear();      
         let newClientSheetLink = newClientSheet.getUrl();                    // Gets the URL of newly copied template file 
         this.set("SheetLink",newClientSheetLink);                            // Places client sheet link to the SheetLink Column
         Dialog.notify("Client Document Structure Created!","Please check column Client Sheet & Client Folder for more details.");
       } else {
          Dialog.notify("Folder not found","Office Use Folder does not exist! Couldn't make a copy of template sheet, please check source folder.");
      }*/
    }
  }

  copyTo(destination) {
    trace(`${this.trace}.copyTo ${destination.trace}`);
    const fields = ["Name", "EmailAddress", "Who"];
    this.copyFieldsTo(destination, fields);
  }

  openClientSheet() {
    trace(`openClientSheet of ${this.trace}`);
    if (!this.isValid) {
      Error.fatal("Please select a valid enquiry.");
    }
    return;
    let sheetId = this.sheetId;
    if (sheetId === "") {
      if (Dialog.confirm("No client sheet found", `There is no client sheet recorded for prospect ${this.name}, create one now?`) == false) {
        return;
      }
      this.createNewClientSheet();
    } else {
      this.clientSheet = null;
      Browser.newTab(this.clientSheet.url);
    }
  }

  draftSelectedEmail() {
    trace(`draftSelectedEmail to ${this.trace}`);
    if (!this.isValid) {
      Error.fatal("Please select a valid enquiry.");
    }
    if (Dialog.confirm("Draft Email", `Draft email to ${this.name}, are you sure?`) == false) {
      return;
    }    
  }  
  
  get name()            { return this.get("Name", "string"); }
  get date()            { return this.get("EventDate"); }
  get fileName()        { return `${Utilities.formatDate(this.date, "GMT+1", "yyyy-MM-dd")} ${this.name}`; }
  get sheetId()         { return this.get("SheetId", "string"); }
  get sheetLink()       { return this.get("SheetLink", "string"); }  
  get folderId()        { return this.get("FolderId", "string"); }
  get folderLink()      { return this.get("FolderLink", "string"); }
  get isValid()         { return this._isValid; }
  get rowOffset()       { return this._rowOffset; }
  set sheetId(value)    { this.set("SheetId", value); }
  set sheetLink(value)  { this.set("SheetLink", value); }
  set folderId(value)   { this.set("FolderId", value); }
  set folderLink(value) { this.set("FolderLink", value); }

  get trace() { return `{Enquiry #${this.rowOffset} ${this._name} ${this.isValid ? "(valid)" : "(invalid)"}}`; }
  
} // Enquiry

//================================================================================================
               
class Prospects {
  
  constructor() {
    this.range = Range.getByName(EnquiriesWaitingResponseRangeName).loadColumnNames();
    trace("NEW " + this.trace);
  }

  get trace() { return `{Prospects ${this.range.trace}}`; }

}


//================================================================================================

class Prospect {
  
  constructor() {
    this.range = Range.getByName(EnquiriesWaitingResponseRangeName).loadColumnNames();
    trace("NEW " + this.trace);
  }

  get trace() { return `{Prospects ${this.range.trace}}`; }

}
