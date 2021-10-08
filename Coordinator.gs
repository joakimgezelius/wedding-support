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


class Coordinator {
  static get eventDetailsRange() { return Range.getByName("EventDetails", "Coordinator").loadColumnNames(); }  
  
  //constructor() {
    //this.range = 
  //}
  
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
    this.eurGbpRate = Range.getByName("EURGBP").value;
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
    ++this.sectionNo;
    this.sectionTitleRange = row.range;
    if (row.itemNo === "" || this.forced) { // Only set item number if empty (or forced)
      row.itemNo = this.generateSectionNo();
    }
  }

  onRow(row) {
    ++this.itemNo;
    trace("EventDetailsUpdater.onRow " + row.itemNo);
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
      row.itemNo = this.generateItemNo();
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

  generateSectionNo() {
    return Utilities.formatString("#%02d", this.sectionNo);
  }

  generateItemNo() {
    if (this.itemNo == 0) 
      return this.generateSectionNo();
    else
      return Utilities.formatString("%s-%02d", this.generateSectionNo(), this.itemNo);
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
    this.eurGbpRate = Range.getByName("EURGBP").value;
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
