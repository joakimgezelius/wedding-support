ClientItineraryRangeName = "ClientItinerary";
ClientItinerarySheetName = "Client Itinerary";

function onUpdateClientItinerary() {
  trace("onUpdateClientItinerary");
  let eventDetails = new EventDetails();
  let clientItineraryBuilder = new ClientItineraryBuilder(Range.getByName(ClientItineraryRangeName, ClientItinerarySheetName));
  eventDetails.sort(SortType.time);
  eventDetails.apply(clientItineraryBuilder);
}


//=============================================================================================
// Class ClientItineraryBuilder
//
class ClientItineraryBuilder {

  constructor(targetRange) {
    this.targetRange = targetRange;
    trace("NEW " + this.trace);
  }

  static formatRange(range) {
    trace("ClientItineraryBuilder.formatRange");
    /*    
    range.setFontWeight("normal")
    .setFontSize(10)
    .breakApart()
    .setBackground("#ffffff")
    .setWrap(true);
    */
  };

  onBegin() {
    trace("ClientItineraryBuilder.onBegin - reset context " + this.trace);
    // Delete all but the first two rows in the target range
    this.targetRange.minimizeAndClear(ClientItineraryBuilder.formatRange);
  }
  
  onEnd() {
    trace("ClientItineraryBuilder.onEnd - trim excess lines");
    this.targetRange.trim();
  }

  onTitle(row) {
    trace("ClientItineraryBuilder.onTitle - ignore row: " + row.title);
  }

  onRow(row) {
    if (row.isItineraryTicked) { // This is an itinerary item
      trace("ClientItineraryBuilder.onRow Ticked: " + row.description);
      let targetRow = this.targetRange.getNextRowAndExtend();
      let column = 1;
      targetRow.getCell(1,column++).setValue(row.date);
      targetRow.getCell(1,column++).setValue(row.time);
      column++; // Gap in the sheet
      targetRow.getCell(1,column++).setValue(row.location);
      targetRow.getCell(1,column++).setValue(row.description);
    } else {
      trace("ClientItineraryBuilder.onRow Unticked, ignore: " + row.description);
    }
  }

  get trace() {
    return `{StaticItineraryBuilder ${this.targetRange.trace}}`;
  }
}
