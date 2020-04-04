
function onUpdateSupplierAccountSummary() {
  trace("onUpdateSupplierAccountSummary");
  var eventDetailsIterator = new EventDetailsIterator();
  var supplierAccountBuilder = new SupplierAccountBuilder("SupplierAccountSummary");
  eventDetailsIterator.iterate(supplierAccountBuilder);
  
}

//=============================================================================================
// Class SupplierAccountBuilder
//
var SupplierAccountBuilder = function(targetRangeName) {
  this.targetRangeName = targetRangeName;
  this.targetRange = Range.getByName(targetRangeName);
  this.targetSheet = this.targetRange.getSheet();
  this.targetRowOffset = 0;
  trace("NEW " + this.trace());
}

SupplierAccountBuilder.prototype.onBegin = function() {
  trace("SupplierAccountBuilder.onBegin - reset context");
}

SupplierAccountBuilder.prototype.onEnd = function() {
  trace("SupplierAccountBuilder.onEnd - fill final title sum & autofit");
}

SupplierAccountBuilder.prototype.onTitle = function(row) {
  this.currentTitle = row.getTitle();
}
  
SupplierAccountBuilder.prototype.onRow = function(row) {
}

SupplierAccountBuilder.prototype.trace = function() {
  return "{SupplierAccountBuilder " + Range.trace(this.targetRange) + "}";
}
