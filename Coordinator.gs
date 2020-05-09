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
    let nativeUnitCostFormula = row.getFormula("NativeUnitCost");
    let currencyFormat = row.currencyFormat;
    //trace(`Cell: ${nativeUnitCostCell.getA1Notation()} ${nativeUnitCostFormula} ${currencyFormat}`);
    budgetUnitCostCell.setNumberFormat(currencyFormat);
    nativeUnitCostCell.setNumberFormat(currencyFormat);
    if (nativeUnitCost === "" || nativeUnitCostFormula !== "") { // Set native unit cost equal to budget unit cost if not set 
      //trace(`Set nativeUnitCost to =${budgetUnitCostA1}`);
      row.nativeUnitCost = `=${budgetUnitCostA1}`;
      nativeUnitCostCell.setFontColor("#ccccff"); // Make text grey
    } else {
      //trace(`setNativeUnitCost: nativeUnitCost="${nativeUnitCost}" nativeUnitCostFormula="${nativeUnitCostFormula}" - make it black`);
      nativeUnitCostCell.setFontColor("#000000"); // Black
    }
  }

  setMarkupAndPrice(row) {
    let unitCostA1 = row.getA1Notation("UnitCost");
    let unitPriceA1 = row.getA1Notation("UnitPrice");
    let markupA1 = row.getA1Notation("Markup");
    let markupFormula = row.getFormula("Markup");
    let unitPriceFormula = row.getFormula("UnitPrice");
    if (markupFormula !== "" && unitPriceFormula !== "") { // Formulas for both markup & unit price, flag it! Keep formula in unit price, set markup to default 30%
      Error.warning(`Found formulas in row ${row.rowPosition}, columns Markup and UnitPrice, only one of the two can be calculated, please choose one to give a value.`);
      row.markup = 0.3;
    }
    if (row.markup === "" && row.unitPrice !== "" && unitPriceFormula === "") { // Only set markup formula if it's empty, and there is a Unit Price set (which is not a formula)
      row.markup = `=IF(OR(${unitCostA1}="", ${unitCostA1}=0, ${unitPriceA1}="", ${unitPriceA1}=0), "", (${unitPriceA1}-${unitCostA1})/${unitCostA1})`;
    } 
    else if (row.unitPrice === "" || this.forced) { // Only set unit price if empty (or forced)
      row.unitPrice = `=IF(OR(${unitCostA1}="", ${unitCostA1}=0), "", ${unitCostA1} * ( 1 + ${markupA1}))`;
    }
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
