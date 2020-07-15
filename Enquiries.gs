enquiriesRangeName = "Enquiries";

enquiriesFolderId = "1EtAGPReyn5ZMyXf6xboCyTncGWsvxNz_";

// Go through raw list of enquiries, 
//  - find those that are work-in-progress
//
function onUpdateEnquiries() {
  trace("onUpdateEnquiries");
  let enquiries = new Enquiries;
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


//================================================================================================

class Enquiries {

  constructor(rangeName = null) {
    rangeName = (rangeName === null) ? enquiriesRangeName : rangeName;
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
    let targetFolder = Folder.getById(EnquiriesFolderId);
    let weddingClientTemplateFile = File.getById(WeddingClientTemplateSpreadsheetId);
    let clientSpreadsheetFile = weddingClientTemplateFile.makeCopy(clientSheetName, targetFolder);
    this.clientSheet = Spreadsheet.openById(clientSpreadsheetFile.id);
    this.sheetId = this.clientSheet.id;
    this.sheetLink = `=hyperlink("${this.clientSheet.url}";"...")`;
    Browser.newTab(this.clientSheet.url);
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
  
  get name()           { return this.get("Name", "string"); }
  get sheetId()        { return this.get("SheetId", "string"); }
  get sheetLink()      { return this.get("SheetLink", "string"); }
  set sheetId(value)   { this.set("SheetId", value); }
  set sheetLink(value) { this.set("SheetLink", value); }
  get isValid()        { return this._isValid; }
  get rowOffset()      { return this._rowOffset; }

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
