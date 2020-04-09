function onTestCase1() {
  trace("onTestCase1");
  EventRow.init();
  let columnLetter = EventRow.columnNumbers.getColumnLetter("Who");
  trace(`EventRow.columnNumbers.getColumnLetter("Who") --> ${columnLetter}`);
//  Error.break;
}

function onTestCase2() {
  
}
