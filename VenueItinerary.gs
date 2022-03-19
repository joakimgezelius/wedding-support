function onUpdateVenueItinerary() {
  trace("onUpdateVenueItinerary");
  let eventDetails = new EventDetails();
  let venueItineraryBuilder = new VenueItineraryBuilder(Range.getByName("VenueItinerary", "Venue Itinerary"));
  trace("onUpdateVenueItinerary: apply VenueItineraryBuilder...");
  //eventDetails.sort(SortType.date);
  eventDetails.apply(venueItineraryBuilder);
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
    range.breakApart().
    setFontWeight("normal").
    setFontSize(10).
    setBackground("#ffffff").
    setWrap(true);
  };

  static formatTitle(range) {
    range.setFontWeight("bold").
    setFontSize(14).
    setBackground("#f2f0ef").
    setHorizontalAlignment("center");
  }
  
  onBegin() {
    trace("VenueItineraryBuilder.onBegin - reset context");
    // Delete all but two lines of the range, clear and set default format
    this.targetRange.minimizeAndClear(VenueItineraryBuilder.formatRange);
    this.currentSection = 0;
    this.sectionItemCount = 0;
  }
  
  onEnd() {
    trace("VenueItineraryBuilder.onEnd - trim excess lines");
    this.targetRange.trim();
  }

  onTitle(row) {
    trace("VenueItineraryBuilder.onTitle " + row.title);
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
    VenueItineraryBuilder.formatTitle(targetRow);
  }

  onRow(row) {
    trace("VenueItineraryBuilder.onRow " + row.title);
    if (row.isSupplierTicked) { // This is a decor summary item
      ++this.sectionItemCount;
      let targetRow = this.targetRange.getNextRowAndExtend();
      let column = 2;
      targetRow.getCell(1,column++).setValue(row.date);
      targetRow.getCell(1,column++).setValue(row.startTime);
      targetRow.getCell(1,column++).setValue(row.endTime);
      targetRow.getCell(1,column++).setValue(row.location);
      targetRow.getCell(1,column++).setValue(row.supplier);
      targetRow.getCell(1,column++).setValue(row.description);
      targetRow.getCell(1,column++).setValue(row.quantity);
    }
    else {
      trace("VenueItineraryBuilder.onRow Unticked, ignore: " + row.description);
    }
  }
  
  get trace() {
    return `{VenueItineraryBuilder ${this.targetRange.trace}}`;
  }
}