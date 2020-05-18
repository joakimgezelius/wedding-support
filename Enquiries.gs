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
    this.range.loadColumnNames();
    trace("NEW " + this.trace);
  }

  static get selected() {
    let enquiries = new Enquiries;
    let selection = enquiries.range.sheet.activeRange;
    if (selection === null) {
      Error.fatal("Please select a valid enquiry.");
    }
    let enquiryRowOffset = selection.rowPosition - enquiries.range.rowPosition;
    trace(`Enquiries.selected ${selection.trace} --> Offset ${enquiryRowOffset}`);
    
    let selectedEnquiry = new Enquiry(enquiries.range, enquiryRowOffset);
    if (!selectedEnquiry.isValid) {
      Error.fatal("Please select a valid enquiry.");
    }
  }
  
  get trace() { return `{Enquiries ${this.range.trace}}`; }
  
} // Enquiries


class Enquiry extends RangeRow {
  
  constructor(enquiriesRange, rowOffset) {
    super(enquiriesRange.values[rowOffset], enquiriesRange.formulas[rowOffset], rowOffset, enquiriesRange);
    this.rowOffset = rowOffset;
    this.name = this.get("Name");
    if (this.name === "") {
      this._isValid = false;
    }
    this._isValid = false;
    trace("NEW " + this.trace);
  }
  
  createNewClientSheet(newName) {
    let weddingClientTemplateSpreadsheet = Spreadsheet.openById(WeddingClientTemplateSpreadsheetId);
    weddingClientTemplateSpreadsheet.copy(newName);
  }
  
  openClientSheet() {
  }
  
  get isValid() { return this._isValid; }
  get trace() { return `{Enquiry #${this.rowOffset} ${this.name}}`; }

} // Enquiry
