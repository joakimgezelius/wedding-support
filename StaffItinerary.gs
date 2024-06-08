function onUpdateStaffItinerary() {
  trace("onUpdateStaffItinerary");
  // Disabled feature
  //let eventDetails = new EventDetails();
  //let itineraryBuilder = new StaffItineraryBuilder(Range.getByName("StaffItinerary", "Staff Itinerary"));
  //eventDetails.sort(SortType.staff);
  //eventDetails.apply(itineraryBuilder);
}


//=============================================================================================
// Class StaffItineraryBuilder
//
class StaffItineraryBuilder {
  
  constructor(targetRange) {
    this.targetRange = targetRange;
    trace("NEW " + this.trace);
  }

  static formatRange(range) {
    trace("StaffItineraryBuilder.formatRange");  
    range.setFontWeight("normal")
    .setFontSize(10)
    .breakApart()
    .setBackground("#ffffff")
    .setWrap(true);
  };
  
  onBegin() {
    trace("StaffItineraryBuilder.onBegin - reset context " + this.trace);
    // Delete all but the first two rows in the target range
    this.targetRange.minimizeAndClear(StaffItineraryBuilder.formatRange);
  }
  
  onEnd() {
    trace("StaffItineraryBuilder.onEnd - trim excess lines");
    this.targetRange.trim();
  }

  onTitle(row) {
    trace("StaffItineraryBuilder.onTitle - ignore row: " + row.title);
  }

  onRow(row) {
    if (row.isStaffTicked) { // This is an itinerary item
      trace("StaffItineraryBuilder.onRow Ticked: " + row.description);
      let targetRow = this.targetRange.getNextRowAndExtend();
      //trace("StaffItineraryBuilder.onRow got target row: " + Range.trace(targetRow));
      let column = 1;
      targetRow.getCell(1,column++).setValue(row.responsible);
      targetRow.getCell(1,column++).setValue(row.date);
      targetRow.getCell(1,column++).setValue(row.time);
      targetRow.getCell(1,column++).setValue(row.endTime);
      targetRow.getCell(1,column++).setValue(row.location);
      targetRow.getCell(1,column++).setValue(row.supplier);
      targetRow.getCell(1,column++).setValue(row.description);
    } else {
      trace("StaffItineraryBuilder.onRow Unticked, ignore: " + row.description);
    }
  }

  get trace() {
    return `{StaffItineraryBuilder ${this.targetRange.trace}}`;
  }
}
