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
  //Class CancelledCoordinatorItems() {
  
  
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
    trace("EventDetailsUpdater.onTitle " + row.title);
    this.itemNo = 0;
    ++this.sectionNo;
    this.sectionTitleRange = row.range;
    if (row.itemNo === "" || this.forced) { // Only set item number if empty (or forced)
      row.itemNo = this.generateSectionNo();
    }
  }

  onRow(row) {
    ++this.itemNo;
    trace("EventDetailsUpdater.onRow " + this.itemNo);
    let currencyA1 = row.getA1Notation("Currency");
    let quantityA1 = row.getA1Notation("Quantity");
    let nativeUnitCostA1 = row.getA1Notation("NativeUnitCost");
    let unitCostA1 = row.getA1Notation("UnitCost");
    let commissionPercentageA1 = row.getA1Notation("CommissionPercentage");
    let unitPriceA1 = row.getA1Notation("UnitPrice");

    if (row.itemNo === "" || this.forced) { // Only set item number if empty (or forced)
      row.itemNo = this.generateItemNo();
    }
    this.setNativeUnitCost(row);    
    row.unitCost = `=IF(OR(${currencyA1}="", ${nativeUnitCostA1}="", ${nativeUnitCostA1}=0), "", IF(${currencyA1}="GBP", ${nativeUnitCostA1}, ${nativeUnitCostA1} / EURGBP))`;
    row.totalCost = `=IF(OR(${quantityA1}="", ${quantityA1}=0, ${unitCostA1}="", ${unitCostA1}=0), "", ${quantityA1} * ${unitCostA1} * (1-${commissionPercentageA1}))`;
    this.setMarkupAndPrice(row);
    row.totalPrice = `=IF(OR(${quantityA1}="", ${quantityA1}=0, ${unitPriceA1}="", ${unitPriceA1}=0), "", ${quantityA1} * ${unitPriceA1})`;
//  row.commission = `=IF(OR(${commissionPercentageA1}="", ${commissionPercentageA1}=0), "", ${quantityA1} * ${unitCostA1} * ${commissionPercentageA1})`;
//  row.quantity (=((hour(K215)*60+minute(K215))-(hour(J215)*60+minute(J215)))/60)
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
    let budgetUnitCostA1 = budgetUnitCostCell.getA1Notation();
    let nativeUnitCost = String(row.nativeUnitCost);
    let nativeUnitCostCell = row.getCell("NativeUnitCost");
    let currencyFormat = row.currencyFormat;
    //trace(`Cell: ${nativeUnitCostCell.getA1Notation()} ${currencyFormat}`);
    budgetUnitCostCell.setNumberFormat(currencyFormat);
    nativeUnitCostCell.setNumberFormat(currencyFormat);
    if (nativeUnitCost === "") { // Set native unit cost equal to budget unit cost if not set 
      //trace(`Set nativeUnitCost to =${budgetUnitCostA1}`);
      row.nativeUnitCost = `=${budgetUnitCostA1}`;
      nativeUnitCostCell.setFontColor("#ccccff"); // Make text grey
    } else {
      //trace(`setNativeUnitCost: nativeUnitCost="${nativeUnitCost} - make it black`);
      nativeUnitCostCell.setFontColor("#000000"); // Black
    }
  }

  setMarkupAndPrice(row) {
    let unitCostA1 = row.getA1Notation("UnitCost");
    let unitPriceA1 = row.getA1Notation("UnitPrice");
    let markupA1 = row.getA1Notation("Markup");
    if (row.unitPrice === "" || this.forced) { // Only set unit price if empty (or forced)
      row.unitPrice = `=IF(OR(${unitCostA1}="", ${unitCostA1}=0), "", ${unitCostA1} * ( 1 + ${markupA1}))`;
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
