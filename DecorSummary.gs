function onUpdateDecorSummary() {
  trace("onUpdateDecorSummary");
  let eventDetailsIterator = new EventDetailsIterator();
  let decorSummaryBuilder = new DecorSummaryBuilder(Range.getByName("DecorSummary", "Decor Summary"));
  eventDetailsIterator.iterate(decorSummaryBuilder);
}


//=============================================================================================
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
      targetRow.getCell(1,column++).setValue(row.links);
      targetRow.getCell(1,column++).setValue(row.notes);
    }
  }
  
  get trace() {
    return `{DecorSummaryBuilder ${this.targetRange.trace}}`;
  }
}
