
function onUpdateSuppliers() {
  trace("onUpdateSuppliers");
  var eventDetailsIterator = new EventDetailsIterator();
  var eventSuppliersBuilder = new EventSuppliersBuilder("EventSuppliers");
  eventDetailsIterator.iterate(eventSuppliersBuilder);
}


//=============================================================================================
// Class EventSuppliersBuilder
//
var EventSuppliersBuilder = function(targetRangeName) {
  this.targetRangeName = targetRangeName;
  this.targetRange = Range.getByName(targetRangeName);
  this.targetRowOffset = 0;
  trace("NEW " + this.trace());
}

EventSuppliersBuilder.prototype.onBegin = function() {
  trace("EventSuppliersBuilder.onBegin - reset context");
  this.targetRowOffset = 0;
  // Delete all but the first and the last row in the target range
  var targetRangeHeight = this.targetRange.getHeight();
  if (targetRangeHeight > 2) {
    this.targetRange.getSheet().deleteRows(this.targetRange.getRowIndex() + 1, targetRangeHeight - 2);
  }
  this.targetRange.setValue("").setFontWeight("normal").setFontSize(10);
}
  
EventSuppliersBuilder.prototype.onEnd = function() {
  trace("EventSuppliersBuilder.onEnd - no-op");
}

EventSuppliersBuilder.prototype.onTitle = function(row) {
  trace("EventSuppliersBuilder.onTitle " + row.getTitle());
  if (row.isSupplierTicked()) { // This is an event Suppliers item
    var targetRow = this.getNextTargetRow();
    targetRow.getCell(1,4).setValue(row.getTitle()).setFontWeight("bold").setFontSize(14);
  }
}

EventSuppliersBuilder.prototype.onRow = function(row) {
  trace("EventSuppliersBuilder.onRow ");
  if (row.isSupplierTicked()) { // This is an event supplier item
    var targetRow = this.getNextTargetRow();
    var column = 1;
    var image = "";
    targetRow.getCell(1,column++).setValue(row.getItemNo());
//  targetRow.getCell(1,column++).setValue(image);
    targetRow.getCell(1,column++).setValue(row.getSupplier());
    targetRow.getCell(1,column++).setValue(row.getStatus());
//  targetRow.getCell(1,column++).setValue(row.getWho());
//  targetRow.getCell(1,column++).setValue(row.getLocation());
    targetRow.getCell(1,column++).setValue(row.getDescription());
    targetRow.getCell(1,column++).setValue(row.getQuantity());

  }
}

// private method getNextTargetRow
//
EventSuppliersBuilder.prototype.getNextTargetRow = function() {
  var targetRow = this.targetRange.offset(this.targetRowOffset++, 0, 1); // A range of 1 row height
  targetRow.getSheet().insertRowAfter(targetRow.getRowIndex());
  return targetRow;
}
  
EventSuppliersBuilder.prototype.trace = function() {
  return "{EventSuppliersBuilder " + Range.trace(this.targetRange) + "}";
}
