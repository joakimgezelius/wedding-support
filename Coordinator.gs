const CleanUpType = { markup: "markup", id: "id" };

function onUpdateCoordinator() {
  trace("onUpdateCoordinator");
  let eventDetails = new EventDetails();
  let eventDetailsUpdater = new EventDetailsUpdater(false);
  eventDetails.apply(eventDetailsUpdater);
}

function onUpdateCoordinatorForced() {
  trace("onUpdateCoordinatorForced");
  if (Dialog.confirm("Forced Coordinator Update - Confirmation Required", "Are you sure you want to force-update the coordinator? It will overwrite row numbers and formulas, make sure the sheet is sorted properly!") == true) {
    let eventDetails = new EventDetails();
    let eventDetailsUpdater = new EventDetailsUpdater(true);
    eventDetails.apply(eventDetailsUpdater);
  }
}

function onReverseMarkupCalculations(){
  trace("onReverseMarkupCalculations");
  if (Dialog.confirm("Reverse Mark-up Calculations - Confirmation Required", "Are you sure you want to reverse the mark-up calculations? It will overwrite mark-up formulas and client unit prices") == true) {
    let eventDetails = new EventDetails();
    let eventDetailsCleaner = new EventDetailsCleaner(CleanUpType.markup);
    eventDetails.apply(eventDetailsCleaner);
  }
}

function onCheckCoordinator() {
  trace("onCheckCoordinator");
  let eventDetails = new EventDetails();
  let eventDetailsChecker = new EventDetailsChecker();
  eventDetails.apply(eventDetailsChecker);
}


//=============================================================================================
// Class EventDetailsUpdater
//

class EventDetailsUpdater {
  
  constructor(forced) {
    this.forced = forced;
    trace("NEW " + this.trace);
  }

  onBegin() {
    this.itemNo = 0;
    this.sectionNo = 0;
    this.eurGbpRate = Spreadsheet.getCellValue("EURGBP");
    trace("EventDetailsUpdater.onBegin - EURGBP=" + this.eurGbpRate);
  }
  
  onEnd() {
    trace("EventDetailsUpdater.onEnd - no-op");
  }

  finalizeSectionFormatting() {
    if (this.sectionNo > 1) {// This is not the first title row

    }
  }
  
  onTitle(row) {
    trace("EventDetailsUpdater.onTitle " + row.itemNo + " " + row.title);
    this.itemNo = 0;
    this.sectionTitleRange = row.range;
    if (row.itemNo === "") { // Only set item id (itemNo) field if empty, if so generate a string on the format #NN   
      this.sectionId = this.generateSectionId();
      row.itemNo = this.sectionId;
    } else {
      this.sectionId = row.itemNo; // Pick up the section ID to use throughout the section
    }
  }

  onRow(row) {
    trace("EventDetailsUpdater.onRow " + row.sectionId + " " + this.itemNo);
    ++this.itemNo;
    let a1_selected = row.getA1Notation("Selected");
    let a1_currency = row.getA1Notation("Currency");
    let a1_quantity = row.getA1Notation("Quantity");
    let a1_nativeUnitCost = row.getA1Notation("NativeUnitCost");
    let a1_vat = row.getA1Notation("VAT");
    let a1_nativeUnitCostWithVAT = row.getA1Notation("NativeUnitCostWithVAT");
    let a1_unitCost = row.getA1Notation("UnitCost");
    let a1_commissionPercentage = row.getA1Notation("CommissionPercentage");
    let a1_markup = row.getA1Notation("Markup");
    let a1_unitPrice = row.getA1Notation("UnitPrice");
    let a1_startTime = row.getA1Notation("Time");
    let a1_endTime = row.getA1Notation("EndTime");    

    if (row.itemNo === "" || this.forced) { // Only set item number if empty (or forced)
      row.itemNo = this.generateItemId();
    }
    this.setNativeUnitCost(row);
    row.nativeUnitCostWithVAT = `=IF(OR(${a1_currency}="", ${a1_nativeUnitCost}="", ${a1_nativeUnitCost}=0), "", ${a1_nativeUnitCost}*(1+${a1_vat}))`;
    row.getCell("NativeUnitCostWithVAT").setNumberFormat(row.currencyFormat);
    row.unitCost = `=IF(OR(${a1_currency}="", ${a1_nativeUnitCostWithVAT}="", ${a1_nativeUnitCostWithVAT}=0), "", IF(${a1_currency}="GBP", ${a1_nativeUnitCostWithVAT}, ${a1_nativeUnitCostWithVAT} / EURGBP))`;
    row.totalGrossCost = `=IF(OR(${a1_quantity}="", ${a1_quantity}=0, ${a1_unitCost}="", ${a1_unitCost}=0, ${a1_selected}=FALSE), "", ${a1_quantity} * ${a1_unitCost})`;
    row.totalNettCost = `=IF(OR(${a1_quantity}="", ${a1_quantity}=0, ${a1_unitCost}="", ${a1_unitCost}=0, ${a1_selected}=FALSE), "", ${a1_quantity} * ${a1_unitCost} * (1-${a1_commissionPercentage}))`;
    row.unitPrice = `=IF(OR(${a1_unitCost}="", ${a1_unitCost}=0), "", ${a1_unitCost} * ( 1 + ${a1_markup}))`;
    row.totalPrice = `=IF(OR(${a1_quantity}="", ${a1_quantity}=0, ${a1_unitPrice}="", ${a1_unitPrice}=0, ${a1_selected}=FALSE), "", ${a1_quantity} * ${a1_unitPrice})`;
//  row.commission = `=IF(OR(${a1_commissionPercentage}="", ${a1_commissionPercentage}=0), "", ${a1_quantity} * ${a1_unitCost} * ${a1_commissionPercentage})`;
    if (row.isStaffTicked && (row.quantity === "")) { // Calculate quantity if staff and time fields are filled 
      row.quantity = `=IF(OR(${a1_startTime}="", ${a1_endTime}=""), "", ((hour(${a1_endTime})*60+minute(${a1_endTime}))-(hour(${a1_startTime})*60+minute(${a1_startTime})))/60)`;
    }
    if (row.isInStock && this.forced) { // Set mark-up and commission for in-stock items
      trace("- Set In Stock commision & mark-up on " + this.itemNo);
      row.commissionPercentage = 0.5;
      row.markup = 0;
    }
  }

  generateSectionId() {
    ++this.sectionNo;
    return Utilities.formatString("#%02d", this.sectionNo);
  }

  generateItemId() {
    return Utilities.formatString("%s-%02d", this.sectionId, this.itemNo);
  }

  setNativeUnitCost(row) {
    let budgetUnitCostCell = row.getCell("BudgetUnitCost");
    let a1_budgetUnitCost = budgetUnitCostCell.getA1Notation();
    let nativeUnitCost = String(row.nativeUnitCost);
    let nativeUnitCostCell = row.getCell("NativeUnitCost");
    let currencyFormat = row.currencyFormat;
    //trace(`Cell: ${nativeUnitCostCell.getA1Notation()} ${currencyFormat}`);
    budgetUnitCostCell.setNumberFormat(currencyFormat);
    nativeUnitCostCell.setNumberFormat(currencyFormat);
    if (nativeUnitCost === "") { // Set native unit cost equal to budget unit cost if not set 
      //trace(`Set nativeUnitCost to =${budgetUnitCostA1}`);
      row.nativeUnitCost = `=${a1_budgetUnitCost}`;
      nativeUnitCostCell.setFontColor("#ccccff"); // Make text grey
    } else {
      //trace(`setNativeUnitCost: nativeUnitCost="${nativeUnitCost} - make it black`);
      nativeUnitCostCell.setFontColor("#000000"); // Black
    }
  }

  get trace() {
    return `{EventDetailsUpdater forced=${this.forced}}`;
  }
}


//=============================================================================================
// Class EventDetailsCleaner
//

class EventDetailsCleaner {
  constructor(type) {
    this.type = type;
    trace("NEW " + this.trace);
  }

  onBegin() {
    this.itemNo = 0;
    this.sectionNo = 0;
    this.eurGbpRate = Spreadsheet.getCellValue("EURGBP");
    trace("EventDetailsCleaner.onBegin - EURGBP=" + this.eurGbpRate);
  }
  
  onEnd() {
    trace("EventDetailsCleaner.onEnd - no-op");
  }

  onTitle(row) {
    trace("EventDetailsCleaner.onTitle " + row.title);
  }

  onRow(row) {
    trace("EventDetailsCleaner.onRow " + row.description);
    switch (this.type) {
      case CleanUpType.markup:
        let markup = row.markup;
        row.markup = markup;
        trace(`EventDetailsCleaner.onRow, set mark-up: ${markup}`);
        let unitCostA1 = row.getA1Notation("UnitCost");
        let markupA1 = row.getA1Notation("Markup");
        let unitPrice = `=IF(OR(${unitCostA1}="", ${unitCostA1}=0), "", ${unitCostA1} * ( 1 + ${markupA1}))`;
        row.unitPrice = unitPrice;
        trace(`EventDetailsCleaner.onRow, set unit price: ${unitPrice}`);
        break;
    }
  }

  get trace() {
    return `{EventDetailsCleaner type=${this.type}}`;
  }
}


//=============================================================================================
// Class EventDetailsChecker
//

class EventDetailsChecker {
  constructor() {
    trace(`NEW ${this.trace}`);
  }
  
  onBegin() {
    trace("EventDetailsChecker.onBegin - no-op");
  }
  
  onEnd() {
    trace("EventDetailsChecker.onEnd - no-op");
  }

  onTitle(row) {
    trace(`EventDetailsChecker.onTitle ${row.title}`);
  }

  onRow(row) {
    trace("EventDetailsChecker.onRow ");
  }
  
  get trace() {
    return "{EventDetailsChecker}";
  }
}

//=============================================================================================
// Class PriceListPackage
//

class PriceListPackage {
  constructor() {
    
  }
}

function onInsertPackage() {
  trace("onInsertPackage");
  let insertionRow = Range.getByName("EventDetailsInsertionRow"); // pick up insertion point range
  trace(`insertionRow: ${insertionRow.trace}`);
  let eventDetails = Range.getByName("EventDetails").loadColumnNames();
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Coordinator"); 
  let packageRowCount = Spreadsheet.getCellValue("SelectedPriceListPackageRowCount"); // pickup number of rows in package
  let categoryRangeName = Spreadsheet.getCellValue("SelectedPriceListCategoryRange"); // pickup price list category
  let packageId = Spreadsheet.getCellValue("SelectedPriceListPackageId");             // pickup price list package id
  trace(`packageRowCount: ${packageRowCount}`);
  // insert packageRowCount lines from the insertionRow down - this will create space enough to paste the saved data 
  // Use https://developers.google.com/apps-script/reference/spreadsheet/sheet#insertrowsbeforebeforeposition,-howmany
  trace(`sheet: ${insertionRow.trace} insert ${packageRowCount} rows before row ${insertionRow.rowPosition}`)
  insertionRow.sheet.nativeSheet.insertRowsBefore(insertionRow.rowPosition, packageRowCount);
  // copy packageRowCount rows from the new insertion row to the old one (the insertion row has now moved down by packageRowCount lines)
  //  - first add rows to the old insertion row to form the destination range for the package, and clear it
  let destinationRange = insertionRow.extend(packageRowCount - 1);
  //destinationRange.clear();
  //  - next pick up the insertion point range after insertion, and add rows to it to match the package data
  let sourceRange = Range.getByName("EventDetailsInsertionRow").extend(packageRowCount - 1); 
  sourceRange.copyTo(destinationRange, SpreadsheetApp.CopyPasteType.PASTE_VALUES);
  // Default formatting for title rows
  eventDetails.forEachRow((range) => {
    const row = new EventRow(range);
    if (row.isTitle) {          
      let rowRange = row.nativeRange.getA1Notation();
      let borderRange = sheet.getRange(rowRange);
      borderRange.setBorder(true, true, true, true, false, false, "#000000", SpreadsheetApp.BorderStyle.DOUBLE);
      row.nativeRange.setFontColor("#434343").setBackground("#d6beb7").setFontSize(12);
    }
  });
  // Default formatting for item rows excluding the title row we get in range & create grouping for found items range by depth 1
  destinationRange.nativeRange.offset(1, 0, packageRowCount - 1).setFontColor("#434343").setBackground("#FFFFFF").setFontSize(9).setWrap(true).breakApart().shiftRowGroupDepth(1);  
  // Alternative - copy directly from the price list, this way we get the formulas, formatting etc
  // let priceListSheet = Spreadsheet.openByUrl(Spreadsheet.getCellValue("PriceListURL"));
  // let categoryRange = priceListSheet.getRangeByName(categoryRangeName); // This gives us the selected category section of the price list
  // let packageFirstRow = categoryRange.values.findIndex(row => row[0] == packageId); // Find the first row with the selected packageId
  // trace(`Package ${packageId}, first row: ${packageFirstRow}`);
  // https://developers.google.com/apps-script/reference/spreadsheet/range#offsetrowoffset,-columnoffset,-numrows
  // let sourceRange = new Range(categoryRange.nativeRange.offset(packageFirstRow, 0, packageRowCount));
  // sourceRange.copyTo(destinationRange, SpreadsheetApp.CopyPasteType.PASTE_VALUES);
  // clears the content if there any by chance in rows below the NONE so template is clean, no data blocks the query to populate
  eventDetails.nativeRange.offset(packageRowCount + 2, 0, sheet.getMaxRows()).clear({contentsOnly: true}); 
  // Finally, set the package selector to "None", so that an active choice is required to continue adding packages
  Range.getByName("SelectedPriceListCategory").value = "None";
  Range.getByName("SelectedPriceListPackage").value = "None";
}
