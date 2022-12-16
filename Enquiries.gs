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

// Open project sheet for the selected project, 
// - create a new sheet if no such sheet exists
//
function onOpenProjectSheet() {
  trace("onOpenProjectSheet");
  let enquiries = new Enquiries;
  enquiries.selected.openProjectSheet();
}

function onCreateNewProjectSheet() {
  trace("onCreateNewProjectSheet");
  let enquiries = new Enquiries;
  enquiries.selected.createNewProjectSheet();
}

function onDraftSelectedEmail() {
  trace("onDraftSelectedEmail");
  let enquiries = new Enquiries;
  enquiries.selected.draftSelectedEmail();
}

function onPrepareProjectStructure() {
  trace("onPrepareProjectStructure");
  let enquiries = new Enquiries;
  let templateFolderLink = Spreadsheet.getCellValueLinkUrl("TemplateProjectFolder");     // W & E's >> Upcoming
  let templateProjectSheetLink = Spreadsheet.getCellValueLinkUrl("TemplateProjectSheet"); // URL to the W & E's template sheet
  enquiries.selected.prepareProjectStructure(templateFolderLink, templateProjectSheetLink);
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
  
  createNewProjectSheet() {
    trace(`createNewProjectSheet for ${this.trace}`);
    if (!this.isValid) {
      Error.fatal("Please select a valid enquiry.");
    };
    let projectSheetName = `${this.name}`;
    //  let targetFolder = Folder.getById(enquiriesFolderId);
    let weddingProjectTemplateFile = File.getById(weddingProjectTemplateSpreadsheetId);
    let projectSpreadsheetFile = weddingProjectTemplateFile.copyTo(targetFolder, projectSheetName);
    this.projectSheet = Spreadsheet.openById(projectSpreadsheetFile.id);
    this.setHyperLink("SheetLink", this.projectSheet.url, "Sheet");
    Browser.newTab(this.projectSheet.url);
  }

  prepareProjectStructure(templateFolderLink, templateProjectSheetLink) {
    trace(`prepareProjectStructure for ${this.trace}`);
    let sourceFolder = Folder.getByUrl(templateFolderLink);
    let destinationFolderURL = Spreadsheet.getCellValueLinkUrl("ProjectFoldersRoot"); // Destination - W & E's param named range ProjectFoldersRoot
    let destinationFolder = Folder.getByUrl(destinationFolderURL);
    let paymentsFoldersRootURL = Spreadsheet.getCellValueLinkUrl("PaymentsFoldersRoot"); // Destination - W & E's param named range PaymentsFoldersRoot
    let paymentsFoldersRoot = Folder.getByUrl(paymentsFoldersRootURL);
    if (destinationFolder.folderExists(this.fileName)) {
      Dialog.notify("Project folder already exists!","Please check the Weddings & Events Folder for more details.");      
    } 
    else {
      Dialog.notify("Preparing the Structure...", "Making the new Project Document Structure, This may take a few seconds...");
      sourceFolder.copyTo(destinationFolder, this.fileName);                    // Copies source folder contents to target folder
      let paymentsFolderName =  this.fileName + " - Payments";
      paymentsFoldersRoot.createFolder(paymentsFolderName);

      let templateSheetFile = File.getByUrl(templateProjectSheetLink);

      let projectFolder = destinationFolder.getSubfolder(this.fileName);         // Gets newly created project folder by name
      let projectFolderLink = projectFolder.url;                                 // Gets the URL of newly created project folder 
      this.setHyperLink("FolderLink",projectFolderLink, "Folder");               // Sets the folder link in the cell in FolderLink Column

      let paymentsFolder = paymentsFoldersRoot.getSubfolder(paymentsFolderName); // Gets newly created payment folder by name
      let paymentsFolderLink = paymentsFolder.url;                               // Returns URL to Payment Folder
      this.setHyperLink("PaymentsLink", paymentsFolderLink, "Payments");         // Sets the URL in the Master sheet

      //let targetFolderName = "Office Use";                                     // Folder name to look for copying the template file in it
      //let subFolder = projectFolder.getSubfolder(targetFolderName);
      //let subFolderId = subFolder.id;                                          // Gets the id of found subfolder "Office Use"

      let targetFolder = projectFolder;                                          // Gets the folder by id to copy the template file in it
      if (targetFolder) {
          projectFolder.createShortcut(paymentsFolder, "Payments");              // Creates shortcut to Payment Folder in Project Folder
          templateSheetFile.copyTo(targetFolder,this.fileName);       
          let newProjectSheet = targetFolder.getFile(this.fileName);             // Gets the newly copied file with given name
          let newProjectSheetId = newProjectSheet.id;                            // Returns the id of found file
          let projectTemplate = Spreadsheet.openById(newProjectSheetId);
          projectTemplate.setActive();                                           // Sets the active spreadsheet Confirmed W & E to new project sheet
          let newProjectSheetLink = newProjectSheet.url;                         // Gets the URL of newly copied template file 
          this.set("SheetLink",newProjectSheetLink);                             // Sets the project sheet link to the cell in SheetLink Column
          Dialog.notify("Project Document Structure Created!","Please check column Project Sheet & Project Folder for more details.");
          Browser.newTab(newProjectSheetLink);
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

  openProjectSheet() {
    trace(`openProjectSheet of ${this.trace}`);
    if (!this.isValid) {
      Error.fatal("Please select a valid project.");
    }
    return;
    let sheetId = this.sheetId;
    if (sheetId === "") {
      if (Dialog.confirm("No project sheet found", `There is no project sheet recorded for prospect ${this.name}, create one now?`) == false) {
        return;
      }
      this.createNewProjectSheet();
    } else {
      this.projectSheet = null;
      Browser.newTab(this.projectSheet.url);
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
  get sheetLink()       { return this.get("SheetLink", "string"); }  
  get folderLink()      { return this.get("FolderLink", "string"); }
  get folderLink()      { return this.get("FolderLink", "string"); }
  get isValid()         { return this._isValid; }
  get rowOffset()       { return this._rowOffset; }
  set sheetLink(value)  { this.set("SheetLink", value); }


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
