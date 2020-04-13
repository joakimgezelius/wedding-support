function onUpdateClientItinerary() {
  trace("onUpdateClientItinerary");
  let eventDetailsIterator = new EventDetailsIterator();
  let clientItineraryBuilder = new ClientItineraryBuilder("ClientItinerary", "Client Itinerary");
  eventDetailsIterator.sort(SortType.time);
  eventDetailsIterator.iterate(clientItineraryBuilder);
}


//=============================================================================================
// Class ClientItineraryBuilder
//
class ClientItineraryBuilder {

  constructor(targetRangeName) {
    this.targetRange = Range.getByName(targetRangeName);
    this.targetRowOffset = 0;
    trace("NEW " + this.trace);
  }

  onBegin() {
    trace("ClientItineraryBuilder.onBegin - reset context " + this.trace);
    // Delete all but the first and the last row in the target range
    this.targetRange.deleteExcessiveRows(2);
    this.targetRange.clear();
    this.targetRowOffset = 0;
  }
  
  onEnd() {
    trace("ClientItineraryBuilder.onEnd - no-op");
  }

  onTitle(row) {
    trace("ClientItineraryBuilder.onTitle - ignore row: " + row.title);
  }

  onRow(row) {
    if (row.isItineraryTicked) { // This is an itinerary item
      trace("ClientItineraryBuilder.onRow Ticked: " + row.description);
      let targetRow = this.getNextTargetRow();
      //trace("StaticItineraryBuilder.onRow got target row: " + Range.trace(targetRow));
      var column = 1;
      targetRow.getCell(1,column++).setValue(row.date);
      targetRow.getCell(1,column++).setValue(row.time);
      targetRow.getCell(1,column++).setValue(row.location);
      targetRow.getCell(1,column++).setValue(row.description);
    } else {
      trace("ClientItineraryBuilder.onRow Unticked, ignore: " + row.description);
    }
  }

  // private method getNextTargetRow
  //
  getNextTargetRow() {
    let targetRow = this.targetRange.range.offset(this.targetRowOffset++, 0, 1); // A range of 1 row height
    targetRow.getSheet().insertRowAfter(targetRow.getRowIndex());
    return targetRow;
  }  
  
  get trace() {
    return `{StaticItineraryBuilder ${this.targetRange.trace}}`;
  }
}
