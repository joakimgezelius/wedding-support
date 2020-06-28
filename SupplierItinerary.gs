function onUpdateSupplierItinerary() {
  trace("onUpdateSupplierItinerary");
  let eventDetails = new EventDetails();
  let itineraryBuilder = new SupplierItineraryBuilder(Range.getByName("SupplierItinerary", "Supplier Itinerary"));
  eventDetails.sort(SortType.supplier);
  eventDetails.apply(itineraryBuilder);
}


//=============================================================================================
// Class SupplierItineraryBuilder
//
class SupplierItineraryBuilder {
  
  constructor(targetRange) {
    this.targetRange = targetRange;
    trace("NEW " + this.trace);
  }

  static formatRange(range) {
    trace("SupplierItineraryBuilder.formatRange");  
    range.setFontWeight("normal")
    .setFontSize(10)
    .breakApart()
    .setBackground("#ffffff")
    .setWrap(true);
  };
  
  onBegin() {
    trace("SupplierItineraryBuilder.onBegin - reset context " + this.trace);
    // Delete all but the first two rows in the target range
    this.targetRange.minimizeAndClear(SupplierItineraryBuilder.formatRange);
  }
  
  onEnd() {
    trace("SupplierItineraryBuilder.onEnd - trim excess lines");
    this.targetRange.trim();
  }

  onTitle(row) {
    trace("SupplierItineraryBuilder.onTitle - ignore row: " + row.title);
  }

  onRow(row) {
    if (row.isSupplierTicked) { // This is an itinerary item
      trace("SupplierItineraryBuilder.onRow Ticked: " + row.description);
      let targetRow = this.targetRange.getNextRowAndExtend();
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

  get trace() {
    return `{SupplierItineraryBuilder ${this.targetRange.trace}}`;
  }
}
