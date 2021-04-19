ClientItineraryRangeName = "ClientItinerary";   
ClientItinerarySheetName = "Client Itinerary";

function onUpdateClientItinerary() {
  trace("onUpdateClientItinerary");
  // Create an EventDetails instance, to iterate over the items in the Corodinator sheet
  let eventDetails = new EventDetails();
  eventDetails.sort(SortType.time);

  // Populate the local Client Itinerary, by applying a ClientItineraryBuilder (defined below) based on the local named range
  let clientItineraryBuilder = new ClientItineraryBuilder(Range.getByName(ClientItineraryRangeName, ClientItinerarySheetName));
  eventDetails.apply(clientItineraryBuilder);

  // 1. Locate named range SharedClientItineraryLink, to pick up the link to the external/shared itinerary sheet
  let sharedClientItineraryLinkCell = Range.getByName("SharedClientItineraryLink", ClientItinerarySheetName);
  // check for an error if there any link is missing then alert
  if (sharedClientItineraryLinkCell == null) {
    Error.fatal("Could not find Named Range SharedClientItineraryLink");
  }

  // 2. Open sheet by using the link in the cell found above (SharedClientItineraryLink)
  let sharedClientItineraryLink = sharedClientItineraryLinkCell.nativeRange.getRichTextValue().getLinkUrl();
  let sharedClientItinerarySheet = Spreadsheet.openByUrl(sharedClientItineraryLink); 

  // 3. Locate the ClientItinerary range in the external/shared, just like the code above
  let sharedClientItineraryRange = sharedClientItinerarySheet.getRangeByName(ClientItineraryRangeName, ClientItinerarySheetName);

  // 4. apply the builder, just as for the local client itinerary
  let sharedClientItineraryBuilder = new ClientItineraryBuilder(sharedClientItineraryRange);  
  eventDetails.apply(sharedClientItineraryBuilder);
} 


//=============================================================================================
// Class ClientItineraryBuilder
//=============================================================================================
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
