const EnquiriesRangeName = "Enquiries";
const ParamsSheetName = "Params";
const CoordinationSheetName = "Coordinator";

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

/* To Prepare Client Document Structure for Small Weddings

function onPrepareClientStructureSmallWedding() {
  trace("onPrepareClientStructureSmallWedding");
  let enquiries = new Enquiries;        
  let sourceFolderId = "1tDi9fOQuQBZSk1z43C6aeo27dAu7WQuS";                                         // Folder ID of Small W & E's template 
  let templateClientSheetLink = Spreadsheet.getCellValueLinkUrl("TemplateClientSheetSmallWedding"); // URL to the Small W & E's template
  enquiries.selected.prepareClientStructure(templateFolderLink, templateClientSheetLink);  
}

// To Prepare Client Document Structure for Large Weddings

function onPrepareClientStructureLargeWedding() {
  trace("onPrepareClientStructureLargeWedding");
  let enquiries = new Enquiries;
  let sourceFolderId = "1hjGms-aGkTqdRtifqXhsF-WPN4uojs2a";                                         // Folder ID of Large W & E's template
  let templateClientSheetLink = Spreadsheet.getCellValueLinkUrl("TemplateClientSheetLargeWedding"); // URL to the Large W & E's template
  enquiries.selected.prepareClientStructure(templateFolderLink, templateClientSheetLink);
}*/

function onPrepareClientStructure() {
  trace("onPrepareClientStructure");
  let enquiries = new Enquiries;
  //let sourceFolderId = "1O_8U4tBeHc4-teA5FBItDi0z1P_FAgtF";                         // Weddings & Events > Templates > Client Folder Template 2022 
  let templateFolderLink = Spreadsheet.getCellValueLinkUrl("TemplateClientFolder");     // W & E's >> Upcoming
  let templateClientSheetLink = Spreadsheet.getCellValueLinkUrl("TemplateClientSheet"); // URL to the W & E's template sheet
  enquiries.selected.prepareClientStructure(templateFolderLink, templateClientSheetLink);
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

  prepareClientStructure(templateFolderLink, templateClientSheetLink) {
    trace(`prepareClientStructure for ${this.trace}`);
    let sourceFolder = Folder.getByUrl(templateFolderLink);
    //let destinationFolderId = "1oHr5tRJJzDq96F8mlHXf2Ikx6aE5KJKf";              // Destination - W & E's >>> Upcoming
    let destinationFolderURL = Spreadsheet.getCellValueLinkUrl("ClientFoldersRoot"); // Destination - W & E's param named range ClientFoldersRoot
    let destinationFolder = Folder.getByUrl(destinationFolderURL);
    let paymentsFoldersRootURL = Spreadsheet.getCellValueLinkUrl("PaymentsFoldersRoot"); // Destination - W & E's param named range PaymentsFoldersRoot
    let paymentsFoldersRoot = Folder.getByUrl(paymentsFoldersRootURL);
    if (destinationFolder.folderExists(this.fileName)) {
      Dialog.notify("Client folder already exists!","Please check the Weddings & Events Folder for more details.");      
    } 
    else {
      Dialog.notify("Preparing the Structure...", "Making the new Client Document Structure, This may take a few seconds...");
      sourceFolder.copyTo(destinationFolder, this.fileName);                    // Copies source folder contents to target folder
      let paymentsFolderName =  this.fileName + " - PAYMENTS";
      sourceFolder.copyTo(paymentsFoldersRoot, paymentsFolderName);

      let templateSheetFile = File.getByUrl(templateClientSheetLink);

      let clientFolder = destinationFolder.getSubfolder(this.fileName);         // Gets newly created client folder by name
      let clientFolderLink = clientFolder.url;                                  // Gets the URL of newly created client folder 
      this.set("FolderLink",clientFolderLink);                                  // Sets the folder link to the cell in FolderLink Column

      let paymentFolder = paymentsFoldersRoot.getSubfolder(paymentsFolderName); // Gets newly created payment folder by name
      let paymentFolderLink = paymentFolder.url;                                // Returns URL to Payment Folder
      this.set("PaymentLink",paymentFolderLink);                                // Sets the URL to the Master sheet

      //let targetFolderName = "Office Use";                                    // Folder name to look for copying the template file in it
      //let subFolder = clientFolder.getSubfolder(targetFolderName);
      //let subFolderId = subFolder.id;                                         // Gets the id of found subfolder "Office Use"

      let targetFolder = clientFolder;                                          // Gets the folder by id to copy the template file in it
      if (targetFolder) {
          let paymentFolderID = paymentFolder.id;                               // Returns ID of Payment Folder of the client
          clientFolder.createShortcut(paymentFolderID);                         // Creates shortcut to Payment Folder in Client Folder
          templateSheetFile.copyTo(targetFolder,this.fileName);       
          let newClientSheet = targetFolder.getFile(this.fileName);              // Gets the newly copied file with given name
          let newClientSheetId = newClientSheet.id;                              // Returns the id of found file
          let clientTemplate = Spreadsheet.openById(newClientSheetId);
          clientTemplate.setActive();                                            // Sets the active spreadsheet Confirmed W & E to new client sheet
          let newClientSheetLink = newClientSheet.url;                           // Gets the URL of newly copied template file 
          this.set("SheetLink",newClientSheetLink);                              // Sets the client sheet link to the cell in SheetLink Column
          Dialog.notify("Client Document Structure Created!","Please check column Client Sheet & Client Folder for more details.");
          Browser.newTab(newClientSheetLink);
       } 
       else {
          Dialog.notify("Folder not found!","Template Folder does not exist! Couldn't make a copy of template sheet, please check source folder.");
      }
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
  // Time-Zone Changes in the Summer Time begins and ends at 1:00 a.m ( Universal Time (GMT))
  get fileName()        { return `${Utilities.formatDate(this.date, "GMT+2", "yyyy-MM-dd")} ${this.name}`; }  
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
