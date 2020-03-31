
//=============================================================================================
// Class Range
//

var Range = function() {
}

Range.clear = function(range) {
  trace("Range.clear " + range);
  range.setValue("");
}

Range.getByName = function(rangeName, spreadsheet) {
  if (spreadsheet == null) {
    spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  }
  var range = spreadsheet.getRangeByName(rangeName)
  if (range === null) {
    Error.fatal("Cannot find named range " + rangeName);
  }
  trace("Range.getByName " + rangeName + " --> " + Range.trace(range));
  return range;
}

Range.trace = function(range) {
  return "{" + range.getSheet().getName() + "!" + range.getA1Notation() + "}";
}

