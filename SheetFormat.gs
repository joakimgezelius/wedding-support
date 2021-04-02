TemplateSpreadsheetId = "1uBFQrefEIyegbogwe8u0r_QYCcD9CxlWh5BMzwBH3bs"; // This is hard-coded to the "2021 Wedding Template" for now

function onApplyFormat() {
  //SpreadsheetApp.getUi().alert("Script is loading");

  // Identify the template (hard coded reference is OK for now)
  let templateSpreadsheet = Spreadsheet.openById(TemplateSpreadsheetId);

  // Get handle to current sheet (tab in spreadsheet)
  let activeSheet = Spreadsheet.active.activeSheet;
  //let mySS = SpreadsheetApp.getActiveSpreadsheet();
  //let sheet = mySS.getActiveSheet();
  //let mainSheet = mySS.getSheetByName("Coordination");
  
  // Get handle to same sheet in the template, get it by looking for the same tab name
  // Iterate over the columns in the current sheet, for each column: {
  let templateSheet = templateSpreadsheet.getSheetByName(activeSheet.name);
  for (let column=1; column <= activeSheet.maxColumns; ++column) {
    trace(`Formatting column ${column}`);
    // Get the column width from the template
    // https://developers.google.com/apps-script/reference/spreadsheet/sheet#getColumnWidth(Integer)
    let width = templateSheet.nativeSheet.getColumnWidth(column);
    // Apply the found column width to the current sheet
    // https://developers.google.com/apps-script/reference/spreadsheet/sheet#setColumnWidth(Integer,Integer)
    activeSheet.nativeSheet.setColumnWidth(column, width);
    trace(`set column with to ${width}`);
  }
}

function onFormatCoordinator() {
  trace("onFormatCoordinator");
  let eventDetails = new EventDetails();
  let eventDetailsFormater = new EventDetailsFormater(false);
  eventDetails.apply(eventDetailsFormater);
}


class EventDetailsFormater {
  constructor(forced) {
    this.forced = forced;
    trace("NEW " + this.trace);
  }

  onBegin() {
    this.itemNo = 0;
    this.sectionNo = 0;
    this.eurGbpRate = Range.getByName("EURGBP").value;
    trace("EventDetailsFormater.onBegin - EURGBP=" + this.eurGbpRate);
  }
  
  onEnd() {
    trace("EventDetailsFormater.onEnd - no-op");
  }

  finalizeSectionFormatting() {
    if (this.sectionNo > 1) {// This is not the first title row

    }
  }
  
  onTitle(row) {
    trace("EventDetailsFormater.onTitle " + row.itemNo + " " + row.title);

  }

  onRow(row) {
    ++this.itemNo;
    trace("EventDetailsFormater.onRow " + row.itemNo);
  }
  
  get trace() {
    return `{EventDetailsFormater forced=${this.forced}}`;
  }
}


