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

  clientSheetList.setQuery("DecorQuery",
    "SELECT '${eventName}',Col1,Col6,Col7,Col11,Col12,Col16,Col17,Col18,Col19,Col20,Col21,Col22,Col23,Col24,Col25,Col26,Col27,Col28,Col29 WHERE Col2=true ORDER BY Col16",
    "SELECT * WHERE Col2<>'#01' ORDER BY Col4,Col5");
}


//=============================================================================================
// Decor Summary tab (in client sheet)
//
function onUpdateDecorSummary() {
  trace("onUpdateDecorSummary");
  let eventDetails = new EventDetails();
  let decorSummaryBuilder = new DecorSummaryBuilder(Range.getByName(DecorSummaryRangeName, DecorSummarySheetName));
  eventDetails.apply(decorSummaryBuilder);
}

//----------------------------------------------------------------------------------------
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
    setBackground("#f2f0ef");
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
//    targetRow.getCell(1,column++).setValue(row.links);
      ++column;
      targetRow.getCell(1,column++).setValue(row.notes);
    }
  }
  
  get trace() {
    return `{DecorSummaryBuilder ${this.targetRange.trace}}`;
  }
}
