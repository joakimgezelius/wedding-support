function onUpdateCoordinator() {
  trace("onUpdateCoordinator");
  if (Dialog.confirm("Update Coordinator - Confirmation Required", "Are you sure you want to update the coordinator? It will overwrite the row numbers, make sure the sheet is sorted properly!") == true) {
    let eventDetailsIterator = new EventDetailsIterator();
    let eventDetailsUpdater = new EventDetailsUpdater();
    eventDetailsIterator.iterate(eventDetailsUpdater);
  }
}

function onCheckCoordinator() {
  trace("onCheckCoordinator");
  let eventDetailsIterator = new EventDetailsIterator();
  let eventDetailsChecker = new EventDetailsChecker();
  eventDetailsIterator.iterate(eventDetailsChecker);
}


//=============================================================================================
// Class EventDetailsUpdater
//

class EventDetailsUpdater {
  constructor() {
    trace("NEW " + this.trace);
  }

  onBegin() {
    this.itemNo = 0;
    this.sectionNo = 0;
    this.eurGbpRate = CRange.getByName("EURGBP").value;
    trace("EventDetailsUpdater.onBegin - EURGBP=" + this.eurGbpRate);
  }
  
  onEnd() {
    trace("EventDetailsUpdater.onEnd - no-op");
  }

  onTitle(row) {
    trace("EventDetailsUpdater.onTitle " + row.title);
    this.itemNo = 0;
    ++this.sectionNo;
    if (row.itemNo === "") { // Only set item number if empty
      row.itemNo = this.generateSectionNo();
    }
  }

  onRow(row) {
    ++this.itemNo;
    trace("EventDetailsUpdater.onRow " + this.itemNo);
    var currencyA1 = row.getA1Notation("Currency");
    var quantityA1 = row.getA1Notation("Quantity");
    var budgetUnitCostA1 = row.getA1Notation("BudgetUnitCost");
    var nativeUnitCostA1 = row.getA1Notation("NativeUnitCost");
    var unitCostA1 = row.getA1Notation("UnitCost");
    var markupA1 = row.getA1Notation("Markup");
    var commissionPercentageA1 = row.getA1Notation("CommissionPercentage");
    var unitPriceA1 = row.getA1Notation("UnitPrice");
    //
    // Set formulas:
    if (row.itemNo === "") { // Only set item number if empty
      row.itemNo = this.generateItemNo();
    }
    if (row.nativeUnitCost === "" || row.nativeUnitCost[0] === "=") { // Set native unit cost equal to budget unit cost if not set
      row.nativeUnitCost = `=${budgetUnitCostA1}`;
    }
    row.unitCost = `=IF(OR(${currencyA1}="", ${nativeUnitCostA1}="", ${nativeUnitCostA1}=0), "", IF(${currencyA1}="GBP", ${nativeUnitCostA1}, ${nativeUnitCostA1} / EURGBP))`;
    row.totalCost = `=IF(OR(${quantityA1}="", ${quantityA1}=0, ${unitCostA1}="", ${unitCostA1}=0), "", ${quantityA1} * ${unitCostA1} * (1-${commissionPercentageA1}))`;
    if (row.markup === "") { // Only set markup if empty
      row.markup = `=IF(OR(${unitCostA1}="", ${unitCostA1}=0, ${unitPriceA1}="", ${unitPriceA1}=0), "", (${unitPriceA1}-${unitCostA1})/${unitCostA1})`;
    }
    if (row.unitPrice === "") { // Only set unit price if empty
      row.unitPrice = `=IF(OR(${unitCostA1}="", ${unitCostA1}=0), "", ${unitCostA1} * ( 1 + ${markupA1}))`;
    }
    row.totalPrice = `=IF(OR(${quantityA1}="", ${quantityA1}=0, ${unitPriceA1}="", ${unitPriceA1}=0), "", ${quantityA1} * ${unitPriceA1})`;
    row.commission = `=IF(OR(${commissionPercentageA1}="", ${commissionPercentageA1}=0), "", ${quantityA1} * ${unitCostA1} * ${commissionPercentageA1})`;
//  if ()
//  row.quantity (=((hour(K215)*60+minute(K215))-(hour(J215)*60+minute(J215)))/60) 
  }

  generateSectionNo() {
    return Utilities.formatString('#%02d', this.sectionNo);
  }

  generateItemNo() {
    if (this.itemNo == 0) 
      return this.generateSectionNo();
    else
      return Utilities.formatString('%s-%02d', this.generateSectionNo(), this.itemNo);
  }

  get trace() {
    return "{EventDetailsUpdater}";
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
