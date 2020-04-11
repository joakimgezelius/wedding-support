function onUpdateSupplierItinerary() {
  trace("onUpdateSupplierItinerary");
  var eventDetailsIterator = new EventDetailsIterator();
  var itineraryBuilder = new SupplierItineraryBuilder("SupplierItinerary", "Supplier Itinerary");
  eventDetailsIterator.sortByTime();
  eventDetailsIterator.iterate(itineraryBuilder);
}


//=============================================================================================
// Class SupplierItineraryBuilder
//
class SupplierItineraryBuilder {
  
  constructor(targetRangeName, targetSheetName) {
    this.targetRange = Range.getByName(targetRangeName, targetSheetName);
    this.targetRowOffset = 0;
    trace("NEW " + this.trace);
  }
  
  onBegin() {
    trace("SupplierItineraryBuilder.onBegin - reset context " + this.trace);
    this.targetRowOffset = 0;
    this.targetRange.deleteExcessiveRows(2); // Keep 2 rows
    this.targetRange.clear();
  }
  
  onEnd() {
    trace("SupplierItineraryBuilder.onEnd - no-op");
  }

  onTitle(row) {
    trace("SupplierItineraryBuilder.onTitle - ignore row: " + row.title);
  }

  onRow(row) {
    if (row.isSupplierTicked) { // This is an itinerary item
      trace("SupplierItineraryBuilder.onRow Ticked: " + row.description);
      let targetRow = this.getNextTargetRow();
      //trace("SupplierItineraryBuilder.onRow got target row: " + Range.trace(targetRow));
      let column = 2;
      targetRow.getCell(1,column++).setValue(row.date);
      targetRow.getCell(1,column++).setValue(row.time);
      targetRow.getCell(1,column++).setValue(row.endTime);
      targetRow.getCell(1,column++).setValue(row.location);
      targetRow.getCell(1,column++).setValue(row.supplier);
      targetRow.getCell(1,column++).setValue(row.description);
    } else {
      trace("SupplierItineraryBuilder.onRow Unticked, ignore: " + row.description);
    }
  }

  // private method getNextTargetRow
  //
  getNextTargetRow() {
    let targetRow = this.targetRange.range.offset(this.targetRowOffset++, 0, 1); // A range of 1 row height
    targetRow.breakApart().setFontWeight("normal").setFontSize(10).setBackground("#ffffff");
//  if (targetRow.getRowIndex() > targetRow.getSheet().getMaxRows()-2) { // We're at the end, extend
    targetRow.getSheet().insertRowAfter(targetRow.getRowIndex());
    return targetRow;
  }

  get trace() {
    return `{SupplierItineraryBuilder ${this.targetRange.trace}}`;
  }
}
