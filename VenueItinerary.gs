function onUpdateVenueItinerary() {
  trace("onUpdateVenueItinerary");
  let eventDetails = new EventDetails();
  let itineraryBuilder = new VenueItineraryBuilder(Range.getByName("VenueItinerary", "Venue Itinerary"));
  eventDetails.sort(SortType.supplier);
  eventDetails.apply(itineraryBuilder);
}


//=============================================================================================
// Class VenueItineraryBuilder
//
class VenueItineraryBuilder {
  
  constructor(targetRange) {
    this.targetRange = targetRange;
    trace("NEW " + this.trace);
  }

  static formatRange(range) {
    trace("VenueItineraryBuilder.formatRange");  
    range.setFontWeight("normal")
    .setFontSize(10)
    .breakApart()
    .setBackground("#ffffff")
    .setWrap(true);
  };
  
  onBegin() {
    trace("VenueItineraryBuilder.onBegin - reset context " + this.trace);
    // Delete all but the first two rows in the target range
    this.targetRange.minimizeAndClear(VenueItineraryBuilder.formatRange);
  }
  
  onEnd() {
    trace("VenueItineraryBuilder.onEnd - trim excess lines");
    this.targetRange.trim();
  }

  onTitle(row) {
    trace("VenueItineraryBuilder.onTitle - ignore row: " + row.title);
  }

  onRow(row) {
    if (row.isVenueTicked) { // This is an itinerary item
      trace("VenueItineraryBuilder.onRow Ticked: " + row.description);
      let targetRow = this.targetRange.getNextRowAndExtend();
      //trace("VenueItineraryBuilder.onRow got target row: " + Range.trace(targetRow));
      let column = 2;
      targetRow.getCell(1,column++).setValue(row.date);
      targetRow.getCell(1,column++).setValue(row.time);
      targetRow.getCell(1,column++).setValue(row.endTime);
      targetRow.getCell(1,column++).setValue(row.location);
      targetRow.getCell(1,column++).setValue(row.supplier);
      targetRow.getCell(1,column++).setValue(row.description);
    } else {
      trace("VenueItineraryBuilder.onRow Unticked, ignore: " + row.description);
    }
  }

  get trace() {
    return `{VenueItineraryBuilder ${this.targetRange.trace}}`;
  }
}
