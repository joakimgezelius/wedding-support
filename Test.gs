function onTestCase1() {
  trace("onTestCase1");
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName("Suppliers Price List");
  let selection = sheet.getSelection();
  let range = selection.getActiveRange();
  Dialog.notify("Selection", range.getA1Notation());
//  Error.break;
}

function onTestCase2() {
  let range = Range.getByName("SupplierItinerary"); //, "Supplier Itinerary");
  trace(range.constructor.name);
  Dialog.notify("Test 2", range.range.constructor.name);
  let test = new Sheet();
}
