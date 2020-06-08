enquiriesRangeName = "Enquiries";
enquiriesNoReplyRangeName = "EnquiriesNoReply";

enquiriesFolderId = "1EtAGPReyn5ZMyXf6xboCyTncGWsvxNz_";


function onUpdateEnquiries() {
  trace("onUpdateEnquiries");
  let enquiries = new Enquiries(enquiriesRangeName);
  let enquiriesNoReply = new Enquiries(enquiriesNoReplyRangeName);
  enquiries.update(enquiriesNoReply);
}

// Open client sheet for the selected client, 
// - create a new sheet if no such sheet exists
//
function onOpenClientSheet() {
  trace("onOpenClientSheet");
  Enquiries.selected.openClientSheet();
}

function onCreateNewClientSheet() {
  trace("onCreateNewClientSheet");
  Enquiries.selected.createNewClientSheet();
}

function onDraftSelectedEmail() {
  trace("onDraftSelectedEmail");
  Enquiries.selected.draftSelectedEmail();
}


//================================================================================================

class Enquiries {

  constructor(rangeName) {
    this.range = Range.getByName(rangeName).loadColumnNames();
    trace("NEW " + this.trace);
  }

  static get selected() {
    let enquiries = new Enquiries;
    let selection = enquiries.range.sheet.activeRange;
    if (enquiries.selection === null) {
      Error.fatal("Please select a valid enquiry.");
    }
    let enquiryRowOffset = selection.rowPosition - enquiries.range.rowPosition;
    trace(`Enquiries.selected ${selection.trace} --> Offset ${enquiryRowOffset}`);

    let enquiryColumnOffset = selection.columnPosition - enquiries.range.rowPosition;
    
    let selectedEnquiry = new Enquiry(enquiries.range, enquiryRowOffset);
    if (!selectedEnquiry.isValid) {
      Error.fatal("Please select a valid enquiry.");
    }
    return selectedEnquiry;
  }
  
  update(target) {
    trace(`${this.trace}.update`);
    for (var rowOffset = 0; rowOffset < this.range.height; rowOffset++) {
      let enquiry = new Enquiry(this.range, rowOffset);
      if (!enquiry.isValid) {
        break;
      }
      target.append(enquiry);
    }
  }
  
  append(enquiry) {
    this.range.findFirstEmptyRow();
    trace(`${this.trace}.append ${enquiry.trace}`);    
//    let targetEnquiry = new Enquiry(this.range, .findFirstEmptyRow();
  }
  
  get trace() { return `{Enquiries ${this.range.trace}}`; }
  
} // Enquiries


//================================================================================================

class Enquiry extends RangeRow {
  
  constructor(enquiriesRange, rowOffset) {
    super(enquiriesRange, rowOffset);
    this.rowOffset = rowOffset;
    this._name = this.name;
    this._isValid = (this._name !== "");
    trace("NEW " + this.trace);
  }
  
  createNewClientSheet() {
    trace(`createNewClientSheet for ${this.trace}`);
    if (!this.isValid) {
      Error.fatal("Please select a valid enquiry.");
    };
    let clientSheetName = `${this._name} (Prospect Client)`;
    let targetFolder = Folder.getById(EnquiriesFolderId);
    let weddingClientTemplateFile = File.getById(WeddingClientTemplateSpreadsheetId);
    let clientSpreadsheetFile = weddingClientTemplateFile.makeCopy(clientSheetName, targetFolder);
    this.clientSheet = Spreadsheet.openById(clientSpreadsheetFile.id);
    this.sheetId = this.clientSheet.id;
    this.sheetLink = `=hyperlink("${this.clientSheet.url}";"...")`;
    Browser.newTab(this.clientSheet.url);
  }
  
  copyTo(destination) {
    const fields = ["nane", "EmailAddress", "Who"];
    this.copyFieldsTo(destination, fields);
  }
  
  openClientSheet() {
    trace(`openClientSheet of ${this.trace}`);
    if (!this.isValid) {
      Error.fatal("Please select a valid enquiry.");
    }
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

  get trace() { return `{Enquiry #${this.rowOffset} ${this.name} ${this.isValid ? "(valid)" : "(invalid)"}}`; }
  
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
