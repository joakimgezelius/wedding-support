
function onUpdateDecorSummary() {
  trace("onUpdateDecorSummary");
  var eventDetailsIterator = new EventDetailsIterator();
  var decorSummaryBuilder = new DecorSummaryBuilder("DecorSummary");
  eventDetailsIterator.iterate(decorSummaryBuilder);
}


//=============================================================================================
// Class DecorSummaryBuilder
//
var DecorSummaryBuilder = function(targetRangeName) {
  this.targetRangeName = targetRangeName;
  this.targetRange = Range.getByName(targetRangeName);
  this.targetRowOffset = 0;
  trace("NEW " + this.trace());
}

DecorSummaryBuilder.prototype.onBegin = function() {
  trace("DecorSummaryBuilder.onBegin - reset context");
  this.targetRowOffset = 0;
  // Delete all but the first and the last row in the target range
  var targetRangeHeight = this.targetRange.getHeight();
  if (targetRangeHeight > 2) {
    this.targetRange.getSheet().deleteRows(this.targetRange.getRowIndex() + 1, targetRangeHeight - 2);
  }
  this.targetRange.setValue("").setFontWeight("normal").setFontSize(10);
}
  
DecorSummaryBuilder.prototype.onEnd = function() {
  trace("DecorSummaryBuilder.onEnd - no-op");
}

DecorSummaryBuilder.prototype.onTitle = function(row) {
  trace("DecorSummaryBuilder.onTitle " + row.getTitle());
  if (row.isDecorTicked()) { // This is a decor summary item
    var targetRow = this.getNextTargetRow();
    targetRow.getCell(1,3).setValue(row.getTitle()).setFontWeight("bold").setFontSize(14);
  }
}

DecorSummaryBuilder.prototype.onRow = function(row) {
  trace("DecorSummaryBuilder.onRow ");
  if (row.isDecorTicked()) { // This is a decor summary item
    var targetRow = this.getNextTargetRow();
    var column = 1;
    var image = "";
    targetRow.getCell(1,column++).setValue(image);
    targetRow.getCell(1,column++).setValue(row.getLocation());
    targetRow.getCell(1,column++).setValue(row.getDescription());
    targetRow.getCell(1,column++).setValue(row.getQuantity());
    targetRow.getCell(1,column++).setValue(row.getUnitPrice());
    targetRow.getCell(1,column++).setValue(row.getTotalPrice());
    targetRow.getCell(1,column++).setValue(row.getClientNotes());
  }
}
  
// private method getNextTargetRow
//
DecorSummaryBuilder.prototype.getNextTargetRow = function() {
  var targetRow = this.targetRange.offset(this.targetRowOffset++, 0, 1); // A range of 1 row height
  targetRow.getSheet().insertRowAfter(targetRow.getRowIndex());
  return targetRow;
}
  
DecorSummaryBuilder.prototype.trace = function() {
  return "{DecorSummaryBuilder " + Range.trace(this.targetRange) + "}";
}
