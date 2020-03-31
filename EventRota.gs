
function onUpdateRota() {
  trace("onUpdateRota");
  var eventDetailsIterator = new EventDetailsIterator();
  var eventRotaBuilder = new EventRotaBuilder("EventRota");  
  eventDetailsIterator.sortByTime();
  eventDetailsIterator.iterate(eventRotaBuilder);
}


//=============================================================================================
// Class EventRotaBuilder
//
var EventRotaBuilder = function(targetRangeName) {
  this.targetRangeName = targetRangeName;
  this.targetRange = Range.getByName(targetRangeName);
  this.targetRowOffset = 0;
  trace("NEW " + this.trace());
}

EventRotaBuilder.prototype.onBegin = function() {
  trace("EventRotaBuilder.onBegin - reset context");
  this.targetRowOffset = 0;
  // Delete all but the first and the last row in the target range
  var targetRangeHeight = this.targetRange.getHeight();
  if (targetRangeHeight > 2) {
    this.targetRange.getSheet().deleteRows(this.targetRange.getRowIndex() + 1, targetRangeHeight - 2);
  }
  this.targetRange.setValue("").setFontWeight("normal").setFontSize(10);
}
  
EventRotaBuilder.prototype.onEnd = function() {
  trace("EventRotaBuilder.onEnd - no-op");
}

EventRotaBuilder.prototype.onTitle = function(row) {
  trace("EventRotaBuilder.onTitle " + row.getTitle() + " - ignore");
//  if (row.isSupplierTicked()) { // This is an event Suppliers item
//    var targetRow = this.getNextTargetRow();
//    targetRow.getCell(1,4).setValue(row.getTitle()).setFontWeight("bold").setFontSize(14);
//  }
}

EventRotaBuilder.prototype.onRow = function(row) {
  trace("EventRotaBuilder.onRow ");
  var who = row.getWho();
  if (who != "") { // This is a rota item (it has somebody assigned)
    var targetRow = this.getNextTargetRow();
    var column = 1;
    var image = "";
//  targetRow.getCell(1,column++).setValue(row.getItemNo());
    targetRow.getCell(1,column++).setValue(image);
    targetRow.getCell(1,column++).setValue(row.getLocation());
    targetRow.getCell(1,column++).setValue(row.getDate());
    targetRow.getCell(1,column++).setValue(row.getTime());
    column++; // End time
    targetRow.getCell(1,column++).setValue(row.getWho());
    targetRow.getCell(1,column++).setValue(row.getDescription());
    targetRow.getCell(1,column++).setValue(row.getQuantity());
    column++; // Delivery check
    column++; // Removal check
    column++; // Delivery location
    targetRow.getCell(1,column++).setValue(row.getItemNotes());
//    targetRow.getCell(1,column++).setValue(row.getSupplier());
//    targetRow.getCell(1,column++).setValue(row.getStatus());
  }
}

// private method getNextTargetRow - extend the target area by inserting a row below the curent one, then return the current row
//
EventRotaBuilder.prototype.getNextTargetRow = function() {
  var targetRow = this.targetRange.offset(this.targetRowOffset++, 0, 1); // A range of 1 row height
  targetRow.getSheet().insertRowAfter(targetRow.getRowIndex());
  return targetRow;
}
  
EventRotaBuilder.prototype.trace = function() {
  return "{EventRotaBuilder " + Range.trace(this.targetRange) + "}";
}
