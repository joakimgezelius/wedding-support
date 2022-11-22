DecorSummaryRangeName = "DecorSummary";
DecorSummarySheetName = "Decor Summary";

//=============================================================================================
// Decor Price List Sheet

function onDecorPriceListPeriodChanged() {
  trace("onDecorPriceListPeriodChanged");
  Dialog.notify("Period Changed", "Sheet will be recalculated, this may take a few seconds...");
  onUpdateDecorPriceList();
  onUpdateNewDecorPriceList();
}

function onUpdateDecorPriceList() {
  trace("onUpdateDecorPriceList");
  let clientSheetList = new ClientSheetList; // In Rota.gs
  let decorPriceList = new DecorPriceList;
  let decorPriceListBuilder = new DecorPriceListBuilder();
  
  clientSheetList.setQuery("DecorQuery",
    "SELECT '${eventName}',Col1,Col6,Col7,Col11,Col12,Col16,Col18,Col20,Col21,Col26,Col27 WHERE Col2=true AND NOT LOWER(Col7) CONTAINS 'cancelled' AND Col16 IS NOT NULL",
    "SELECT * WHERE Col2<>'#01' AND Col9 IS NOT NULL ORDER BY Col7"); 

  decorPriceList.apply(decorPriceListBuilder); 
}

function onUpdateNewDecorPriceList() {
  trace("onUpdateNewDecorPriceList");
  let clientSheetList = new ClientSheetList;  
  let decorPriceList = new DecorPriceList;
  let decorPriceListBuilder = new DecorPriceListBuilder();
 
  clientSheetList.setQuery("NewDecorQuery",
    "SELECT '${eventName}',Col1,Col6,Col7,Col11,Col12,Col16,Col19,Col21,Col22,Col27,Col28 WHERE Col2=true AND NOT LOWER(Col7) CONTAINS 'cancelled' AND Col16 IS NOT NULL",
    "SELECT * WHERE Col2<>'#01' AND Col9 IS NOT NULL ORDER BY Col7");   
  
  decorPriceList.apply(decorPriceListBuilder);
}

//----------------------------------------------------------------------------------------------
// Decor Price List Builder
//

class DecorPriceListBuilder {
  
  constructor(forced) {
    this.forced = forced;
    trace("NEW " + this.trace);
  }

  /*static formatRange(range) {
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
  }*/

  onEnd() {
    trace("DecorPriceListBuilder.onEnd - trim excess lines");
  }

  onRow(row) {
    trace("DecorPriceListBuilder.onRow " +  row.itemNo);
    let a1_currency = row.getA1Notation("Currency");
    let a1_nativeUnitCost = row.getA1Notation("NativeUnitCost");
    let a1_vat = row.getA1Notation("VAT");
    let a1_nativeUnitCostWithVAT = row.getA1Notation("NativeUnitCostWithVAT");
    let a1_unitCost = row.getA1Notation("UnitCost");
    let a1_markup = row.getA1Notation("Markup");    
    //let a1_commissionPercentage = row.getA1Notation("CommissionPercentage");
    //let a1_unitPrice = row.getA1Notation("UnitPrice");

    row.nativeUnitCostWithVAT = `=IF(OR(${a1_currency}="", ${a1_nativeUnitCost}="", ${a1_nativeUnitCost}=0), "", ${a1_nativeUnitCost}*(1+${a1_vat}))`;
    row.getCell("NativeUnitCostWithVAT").setNumberFormat(row.currencyFormat);
    row.unitCost = `=IF(OR(${a1_currency}="", ${a1_nativeUnitCostWithVAT}="", ${a1_nativeUnitCostWithVAT}=0), "", IF(${a1_currency}="GBP", ${a1_nativeUnitCostWithVAT}, ${a1_nativeUnitCostWithVAT} / EURGBP))`;
    row.unitPrice = `=IF(OR(${a1_unitCost}="", ${a1_unitCost}=0), "", ${a1_unitCost} * ( 1 + ${a1_markup}))`;

  }

  static get decorPriceListRange() { 
    return Range.getByName("DecorPriceListNew", "Decor Price List (NEW)").loadColumnNames(); 
  }
  
  get trace() {
    return `{DecorPriceListBuilder forced=${this.forced}}`;
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
      targetRow.getCell(1,column++).setValue(row.incharge);
      targetRow.getCell(1,column++).setValue(row.location);
      targetRow.getCell(1,column++).setValue(row.description);
      targetRow.getCell(1,column++).setValue(row.quantity);
      targetRow.getCell(1,column++).setValue(row.status);      
//    targetRow.getCell(1,column++).setValue(row.links);
      targetRow.getCell(1,column++).setValue(row.notes);
    }
  }
  
  get trace() {
    return `{DecorSummaryBuilder ${this.targetRange.trace}}`;
  }
}

