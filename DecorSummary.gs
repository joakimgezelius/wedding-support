DecorSummaryRangeName = "DecorSummary";
DecorSummarySheetName = "Decor Summary";

//=============================================================================================
// Decor Price List Sheet

function onDecorPriceListPeriodChanged() {
  trace("onDecorPriceListPeriodChanged");
  Dialog.notify("Period Changed", "Sheet will be recalculated, this may take a few seconds...");
  onUpdateDecorPriceList();
}

function onUpdateDecorPriceList() {
  trace("onUpdateDecorPriceList");
  let clientSheetList = new ClientSheetList; // In Rota.gs
  let eventDetails = new EventDetails;
  let decorPriceListBuilder = new DecorPriceListBuilder(Range.getByName("DecorPriceList", "Decor Price List"));

  clientSheetList.setQuery("DecorQuery",
    "SELECT '${eventName}',Col1,Col6,Col7,Col11,Col12,Col16,Col18,Col20,Col21,Col22,Col23,Col26,Col27,Col28 WHERE Col2=true AND NOT LOWER(Col7) CONTAINS 'cancelled' AND Col16 IS NOT NULL",
    "SELECT * WHERE Col2<>'#01' AND Col9 IS NOT NULL ORDER BY Col7");  
  
  //trace("onUpdateDecorPriceList: apply DecorPriceListBuilder...");
  //eventDetails.apply(decorPriceListBuilder);
}

//----------------------------------------------------------------------------------------------
// Decor Price List Builder

class DecorPriceListBuilder {
  
  constructor(targetRange) {
    this.targetRange = targetRange;
    trace("NEW " + this.trace);
  }

  static formatRange(range) {
    trace("DecorPriceListBuilder.formatRange");
    range.breakApart().
    setFontWeight("normal").
    setFontSize(10).
    setBackground("#ffffff").
    setWrap(true);
  };

  static formatTitle(range) {
    range.setFontWeight("bold").
    setFontSize(14).
    setBackground("#f2f0ef");
  }
  
  onBegin() {
    trace("DecorPriceListBuilder.onBegin - reset context");
    // Delete all but two lines of the range, clear and set default format
    //this.targetRange.minimizeAndClear(DecorPriceListBuilder.formatRange);
    this.currentSection = 0;
    this.sectionItemCount = 0;
  }
  
  onEnd() {
    trace("DecorPriceListBuilder.onEnd - trim excess lines");
    this.targetRange.trim();
  }

  onTitle(row) {
    trace("DecorPriceListBuilder.onTitle " + row.title);
    ++this.currentSection;
    if (this.currentSection > 1) { // This is not the first section (if it is, no need for clean-up house-keeping)
      if (this.sectionItemCount == 0) {         // Section with no content?
        this.targetRange.getPreviousRow();      //  - back up instead of moving on
      }
      else {
        this.targetRange.getNextRowAndExtend(); //  - else leave one blank row before next section
      }
    }
    this.sectionItemCount = 0;
    let targetRow = this.targetRange.getNextRowAndExtend(); 
    targetRow.merge();
    targetRow.getCell(1,1).setValue(row.title);
    DecorPriceListBuilder.formatTitle(targetRow);
  }

  onRow(row) {
    trace("DecorPriceListBuilder.onRow " + row.title);
    if (row.isDecorTicked) { // This is a decor summary item
      ++this.sectionItemCount;
      let targetRow = this.targetRange.getNextRowAndExtend();
      let column = 1;
      targetRow.getCell(1,column++).setValue(row.eventName);
      targetRow.getCell(1,column++).setValue(row.itemNo);
      targetRow.getCell(1,column++).setValue(row.category);
      targetRow.getCell(1,column++).setValue(row.status);
      targetRow.getCell(1,column++).setValue(row.location);
      targetRow.getCell(1,column++).setValue(row.supplier);
      targetRow.getCell(1,column++).setValue(row.description);
      targetRow.getCell(1,column++).setValue(row.currency);
      targetRow.getCell(1,column++).setValue(row.nativeUnitCost);
      targetRow.getCell(1,column++).setValue(row.nativeUnitCostWithVAT);
      targetRow.getCell(1,column++).setValue(row.unitCost);  
      targetRow.getCell(1,column++).setValue(row.totalGrossCost);
      targetRow.getCell(1,column++).setValue(row.markup);
      targetRow.getCell(1,column++).setValue(row.commisionPercentage);
      targetRow.getCell(1,column++).setValue(row.unitPrice);
      ++column;
    }
  }
  
  get trace() {
    return `{DecorPriceListBuilder ${this.targetRange.trace}}`;
  }
}

//=============================================================================================
// Decor Summary tab (in client sheet)
//
function onUpdateDecorSummary() {
  trace("onUpdateDecorSummary");
  let eventDetails = new EventDetails;
  let decorSummaryBuilder = new DecorSummaryBuilder(Range.getByName(DecorSummaryRangeName, DecorSummarySheetName));
  trace("onUpdateDecorSummary: apply decorSummaryBuilder...");
  eventDetails.apply(decorSummaryBuilder);
}

//--------------------------------------------------------------------------------------------
// Class DecorSummaryBuilder
//
class DecorSummaryBuilder {
  
  constructor(targetRange) {
    this.targetRange = targetRange;
    trace("NEW " + this.trace);
  }

  static formatRange(range) {
    trace("DecorSummaryBuilder.formatRange");
    range.breakApart().
    setFontWeight("normal").
    setFontSize(10).
    setBackground("#ffffff").
    setWrap(true);
  };

  static formatTitle(range) {
    range.setFontWeight("bold").
    setFontSize(14).
    setBackground("#f2f0ef").
    setHorizontalAlignment("center");
  }
  
  onBegin() {
    trace("DecorSummaryBuilder.onBegin - reset context");
    // Delete all but two lines of the range, clear and set default format
    this.targetRange.minimizeAndClear(DecorSummaryBuilder.formatRange);
    this.currentSection = 0;
    this.sectionItemCount = 0;
  }
  
  onEnd() {
    trace("DecorSummaryBuilder.onEnd - trim excess lines");
    this.targetRange.trim();
  }

  onTitle(row) {
    trace("DecorSummaryBuilder.onTitle " + row.title);
    ++this.currentSection;
    if (this.currentSection > 1) { // This is not the first section (if it is, no need for clean-up house-keeping)
      if (this.sectionItemCount == 0) {         // Section with no content?
        this.targetRange.getPreviousRow();      //  - back up instead of moving on
      }
      else {
        this.targetRange.getNextRowAndExtend(); //  - else leave one blank row before next section
      }
    }
    this.sectionItemCount = 0;
    let targetRow = this.targetRange.getNextRowAndExtend(); 
    targetRow.merge();
    targetRow.getCell(1,1).setValue(row.title);
    DecorSummaryBuilder.formatTitle(targetRow);
  }

  onRow(row) {
    trace("DecorSummaryBuilder.onRow " + row.title);
    if (row.isDecorTicked) { // This is a decor summary item
      ++this.sectionItemCount;
      let targetRow = this.targetRange.getNextRowAndExtend();
      let column = 1;
      let image = "";
      let itemStoreLocation = "";
//    targetRow.getCell(1,column++).setValue(image);
      targetRow.getCell(1,column++).setValue(itemStoreLocation);
      targetRow.getCell(1,column++).setValue(row.supplier);
      targetRow.getCell(1,column++).setValue(row.date);
      targetRow.getCell(1,column++).setValue(row.startTime);
      targetRow.getCell(1,column++).setValue(row.endTime);
      targetRow.getCell(1,column++).setValue(row.who);
      targetRow.getCell(1,column++).setValue(row.location);
      targetRow.getCell(1,column++).setValue(row.description);
      targetRow.getCell(1,column++).setValue(row.quantity);
      targetRow.getCell(1,column++).setValue(row.status);      
//    targetRow.getCell(1,column++).setValue(row.links);
      ++column;
      targetRow.getCell(1,column++).setValue(row.notes);
    }
  }
  
  get trace() {
    return `{DecorSummaryBuilder ${this.targetRange.trace}}`;
  }
}

