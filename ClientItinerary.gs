const ClientItineraryRangeName = "ClientItinerary";
const ClientItinerarySheetName = "Client Itinerary";

function onUpdateClientItinerary() {
  trace("onUpdateClientItinerary");
  // Create an EventDetails instance, to iterate over the items in the Corodinator sheet
  let eventDetails = new EventDetails();
  eventDetails.sort(SortType.time);

  // Populate the local Client Itinerary, by applying a ClientItineraryBuilder (defined below) based on the local named range
  let clientItineraryBuilder = new ClientItineraryBuilder(Range.getByName(ClientItineraryRangeName, ClientItinerarySheetName));
  eventDetails.apply(clientItineraryBuilder);

  // Locate named range SharedClientItineraryLink, to pick up the link to the external/shared Client Itinerary spreadsheet
  let sharedClientItineraryLinkCell = Range.getByName("SharedClientItineraryLink", ClientItinerarySheetName);
  // If named range isn't found then alert and abort
  if (sharedClientItineraryLinkCell == null) {
    Error.fatal("Could not find Named Range SharedClientItineraryLink");
  }

  // Open shared/external Client Itinerary spreadsheet by using the link in the cell SharedClientItineraryLink
  let sharedClientItineraryLink = sharedClientItineraryLinkCell.nativeRange.getRichTextValue().getLinkUrl();
  let sharedClientItinerarySheet = Spreadsheet.openByUrl(sharedClientItineraryLink);

  // Locate the ClientItinerary range in the external/shared client itinerary
  let sharedClientItineraryRange = sharedClientItinerarySheet.getRangeByName(ClientItineraryRangeName, ClientItinerarySheetName)

  // Apply the builder, just as for the local client itinerary
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
