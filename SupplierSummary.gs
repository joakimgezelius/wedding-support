
function onUpdateSupplierSummary() {
  trace("onUpdateSupplierSummary");
  var eventDetailsIterator = new EventDetailsIterator();
  var supplierSummaryBuilder = new SupplierSummaryBuilder("SupplierSummary");
  eventDetailsIterator.iterate(supplierSummaryBuilder);
  
}

//=============================================================================================
// Class SupplierSummaryBuilder
//
var SupplierSummaryBuilder = function(targetRangeName) {
  this.targetRange = Range.getByName(targetRangeName);
  this.targetSheet = this.targetRange.getSheet();
  this.targetRowOffset = 0;
  this.maxTargetRows = 1000;
  trace("NEW " + this.trace());
}

SupplierSummaryBuilder.prototype.onBegin = function() {
  trace("SupplierSummaryBuilder.onBegin - reset context");
  this.currentTitle = "";
  this.currentTitleSumCell = null;
  this.currentTitleSum = 0;
  Range.clear(this.targetRange);
  var column3 = this.targetRange.offset(0, 2, this.targetRange.getNumRows(), 1); // Select a full column
  var column4 = column3.offset(0, 1);
  var column5 = column4.offset(0, 1);
  column3.setNumberFormat("0");
  column4.setNumberFormat("£#,##0.00");
  column5.setNumberFormat("£#,##0.00");
  this.targetRange.setFontWeight("normal").setFontSize(10);
}

SupplierSummaryBuilder.prototype.onEnd = function() {
  trace("SupplierSummaryBuilder.onEnd - fill final title sum & autofit");
  this.fillTitleSum(null);
  this.targetSheet.setColumnWidth(1, 1);
  this.targetSheet.setColumnWidth(2, 500);
//  this.targetSheet.autoResizeColumns(3, this.targetSheet.getMaxColumns());
}

SupplierSummaryBuilder.prototype.onTitle = function(row) {
  this.currentTitle = row.getTitle();
  trace("SupplierSummaryBuilder.onTitle " + this.currentTitle);
  this.getNextTargetRow(); // Leave one blank row
  var targetRow = this.getNextTargetRow();
  var column = 1;
  targetRow.getCell(1,column++).setValue(this.currentTitle);
  targetRow.getCell(1,column++).setValue(this.currentTitle);
  ++column;
  ++column;
  this.fillTitleSum(targetRow.getCell(1,column++));
  targetRow.setFontWeight("bold");
  targetRow.setFontSize(12);
}

SupplierSummaryBuilder.prototype.onRow = function(row) {
  var totalPrice = row.getTotalPrice();
  if (totalPrice > 0) {
    trace("SupplierSummaryBuilder.onRow " + row.getDescription());
    var targetRow = this.getNextTargetRow();
    var column = 1;
    targetRow.getCell(1,column++).setValue(this.currentTitle);
    targetRow.getCell(1,column++).setValue(row.getDescription());
    targetRow.getCell(1,column++).setValue(row.getQuantity());
    targetRow.getCell(1,column++).setValue(row.getUnitPrice());
    targetRow.getCell(1,column++).setValue(totalPrice);
    targetRow.setFontWeight("normal");
    targetRow.setFontSize(10);
    this.currentTitleSum += totalPrice;
  } else {
    trace("SupplierSummaryBuilder.onRow - ignore (no price): " + row.getDescription());
  }
}

SupplierSummaryBuilder.prototype.fillTitleSum = function(nextTitleSumCell) {
  if (this.currentTitleSumCell != null) {
    this.currentTitleSumCell.setValue(this.currentTitleSum);
    this.currentTitleSum = 0;
  }
  this.currentTitleSumCell = nextTitleSumCell;
}

// private method getNextTargetRow
//
SupplierSummaryBuilder.prototype.getNextTargetRow = function() {
  return this.targetRange.offset(Math.min(this.targetRowOffset++, this.maxTargetRows-1), 0, 1); // A range of 1 row height
}

SupplierSummaryBuilder.prototype.trace = function() {
  return "{SupplierSummaryBuilder " + Range.trace(this.targetRange) + "}";
}
