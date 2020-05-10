EnquiriesRangeName = "Enquiries";

function onUpdateEnquiries() {
  trace("onUpdateEnquiries");
}

function onOpenClientSheet() {
  trace("onOpenClientSheet");
  Enquiries.selected.openClientSheet();
}

function onCreateNewClientSheet() {
  trace("onCreateNewClientSheet");
  Enquiries.selected.createNewClientSheet("testing");
}


class Enquiries {

  constructor() {
    this.range = Range.getByName(EnquiriesRangeName);
    trace("NEW " + this.trace);
  }

  static get selected() {
    let enquiries = new Enquiries;
    let selection = enquiries.range.sheet.activeRange;
    console.log(selection);
    if (selection === null) {
      Error.fatal("Please select an enquiry.");
    }
    
    trace(`Enquiries.selected ${selection.trace}`);

    return new Enquiry();
  }
  
  get trace() { return `{Enquiries ${this.range.trace}}`; }
  
} // Enquiries


class Enquiry {
  
  constructor() {
  }
  
  createNewClientSheet(newName) {
    let weddingClientTemplateSpreadsheet = Spreadsheet.openById(WeddingClientTemplateSpreadsheetId);
    weddingClientTemplateSpreadsheet.copy(newName);
  }
  
  openClientSheet() {
  }
  
} // Enquiry
