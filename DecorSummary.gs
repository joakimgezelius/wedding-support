
function onUpdateDecorSummary() {
  trace("onUpdateDecorSummary");
  let eventDetailsIterator = new EventDetailsIterator();
  let decorSummaryBuilder = new DecorSummaryBuilder("DecorSummary");
  eventDetailsIterator.iterate(decorSummaryBuilder);
}


//=============================================================================================
// Class DecorSummaryBuilder
//
class DecorSummaryBuilder {
  
  constructor(targetRangeName) {
    this.targetRange = Range.getByName(targetRangeName);
    this.targetRowOffset = 0;
    trace("NEW " + this.trace);
  }

  onBegin() {
    trace("DecorSummaryBuilder.onBegin - reset context");
    this.targetRowOffset = 0;
    // Delete all but the first and the last row in the target range
    this.targetRange.deleteExcessiveRows(2); // Keep 2 rows
    this.targetRange.clear();
    this.targetRange.range.breakApart().setFontWeight("normal").setFontSize(10).setBackground("#ffffff").setWrap(true);
  }
  
  onEnd() {
    trace("DecorSummaryBuilder.onEnd - no-op");
  }

  onTitle(row) {
    trace("DecorSummaryBuilder.onTitle " + row.title);
    if (row.isDecorTicked) { // This is a decor summary item
      let targetRow = this.getNextTargetRow();
      targetRow.merge();
      targetRow.getCell(1,1).setValue(row.title).setFontWeight("bold").setFontSize(14).setBackground("#f2f0ef");
    }
  }

  onRow(row) {
    trace("DecorSummaryBuilder.onRow");
    if (row.isDecorTicked) { // This is a decor summary item
      let targetRow = this.getNextTargetRow();
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
  
  // private method getNextTargetRow
  //
  getNextTargetRow() {
    let targetRow = this.targetRange.range.offset(this.targetRowOffset++, 0, 1); // A range of 1 row height
    targetRow.getSheet().insertRowAfter(targetRow.getRowIndex());
    targetRow.breakApart().setFontWeight("normal").setFontSize(10).setBackground("#ffffff");
    return targetRow;
  }
  
  get trace() {
    return `{DecorSummaryBuilder ${this.targetRange.trace}}`;
  }
}
