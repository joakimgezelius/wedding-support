function onTestCase1() {
  trace("onTestCase1");
  EventRow.init();
  let columnLetter = EventRow.columnNumbers.getColumnLetter("Who");
  trace(`EventRow.columnNumbers.getColumnLetter("Who") --> ${columnLetter}`);
//  Error.break;
}

function onTestCase2() {
  let range = Range.getByName("SupplierItinerary"); //, "Supplier Itinerary");
  trace(range.constructor.name);
  Dialog.notify("Test 2", range.range.constructor.name);
  let test = new Sheet();
}
