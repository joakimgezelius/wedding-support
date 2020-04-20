function onUpdateCoordinator() {
  trace("onUpdateCoordinator");
  let eventDetailsIterator = new EventDetailsIterator();
  let eventDetailsUpdater = new EventDetailsUpdater(false);
}

function onUpdateCoordinatorForced() {
  trace("onUpdateCoordinatorForced");
  if (Dialog.confirm("Forced Coordinator Update - Confirmation Required", "Are you sure you want to force-update the coordinator? It will overwrite row numbers and formulas, make sure the sheet is sorted properly!") == true) {
    let eventDetailsIterator = new EventDetailsIterator();
    let eventDetailsUpdater = new EventDetailsUpdater(true);
    eventDetailsIterator.iterate(eventDetailsUpdater);
  }
}

function onCheckCoordinator() {
  trace("onCheckCoordinator");
  let eventDetailsIterator = new EventDetailsIterator();
  let eventDetailsChecker = new EventDetailsChecker();
  eventDetailsIterator.iterate(eventDetailsChecker);
}


class Coordinator {
  static get eventDetailsRange() { return Range.getByName("EventDetails", "Coordinator"); }
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
    if (row.itemNo === "" || this.forced) { // Only set item number if empty (or forced)
      row.itemNo = this.generateItemNo();
    }
    let nativeUnitCost = String(row.nativeUnitCost);
    let nativeCurrencyFormat = (row.currency === "GBP") ? "£#,##0" : "€#,##0";
    let nativeUnitCostCell = row.getCell("NativeUnitCost").setNumberFormat(nativeCurrencyFormat);
    if (nativeUnitCost === "" || nativeUnitCost.charAt(0) === "=") { // Set native unit cost equal to budget unit cost if not set
      row.nativeUnitCost = `=${budgetUnitCostA1}`;
      nativeUnitCostCell.setFontColor("#aaaaaa"); // Format
    } else {
      nativeUnitCostCell.setFontColor("#000000"); // Black
    }
    
    row.unitCost = `=IF(OR(${currencyA1}="", ${nativeUnitCostA1}="", ${nativeUnitCostA1}=0), "", IF(${currencyA1}="GBP", ${nativeUnitCostA1}, ${nativeUnitCostA1} / EURGBP))`;
    row.totalCost = `=IF(OR(${quantityA1}="", ${quantityA1}=0, ${unitCostA1}="", ${unitCostA1}=0), "", ${quantityA1} * ${unitCostA1} * (1-${commissionPercentageA1}))`;
    if (row.markup === "") { // Only set markup if empty
      row.markup = `=IF(OR(${unitCostA1}="", ${unitCostA1}=0, ${unitPriceA1}="", ${unitPriceA1}=0), "", (${unitPriceA1}-${unitCostA1})/${unitCostA1})`;
    }
    if (row.unitPrice === "" || this.forced) { // Only set unit price if empty (or forced)
      row.unitPrice = `=IF(OR(${unitCostA1}="", ${unitCostA1}=0), "", ${unitCostA1} * ( 1 + ${markupA1}))`;
    }
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

  get trace() {
    return `{EventDetailsUpdater forced=${this.forced}}`;
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
