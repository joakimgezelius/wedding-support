
function onUpdateBudget() {
  trace("onUpdateBudget");
  var eventDetailsIterator = new EventDetailsIterator();
  var budgetBuilder = new BudgetBuilder("Budget");
  eventDetailsIterator.iterate(budgetBuilder);
}

//=============================================================================================
// Class BudgetBuilder
//
var BudgetBuilder = function(targetRangeName) {
  this.targetRangeName = targetRangeName;
  this.targetRange = Range.getByName(targetRangeName);
  this.targetSheet = this.targetRange.getSheet();
  this.targetRowOffset = 0;
  trace("NEW " + this.trace());
}

BudgetBuilder.prototype.onBegin = function() {
  trace("BudgetBuilder.onBegin - reset context");
  this.currentTitle = "";
  this.currentTitleSumCell = null;
  this.currentTitleSum = 0;
  Range.clear(this.targetRange);
  this.targetRange.setFontWeight("normal").setFontSize(10).breakApart().setBackground("#ffffff").setWrap(true);
  var targetRangeHeight = this.targetRange.getHeight();
  if (targetRangeHeight > 2) {
    this.targetRange.getSheet().deleteRows(this.targetRange.getRowIndex() + 1, targetRangeHeight - 3);
    this.targetRange = Range.getByName(this.targetRangeName); // Reload the range as it has changed now
  }
  trace("BudgetBuilder.onBegin - format columns");
  var column3 = this.targetRange.offset(0, 2, this.targetRange.getNumRows(), 1); // Select a full column
  var column4 = column3.offset(0, 1);
  var column5 = column4.offset(0, 1);
  column3.setNumberFormat("0");
  column4.setNumberFormat("£#,##0.00");
  column5.setNumberFormat("£#,##0.00");
  trace("BudgetBuilder.onBegin - done");
}

BudgetBuilder.prototype.onEnd = function() {
  trace("BudgetBuilder.onEnd - fill final title sum & autofit");
  this.fillTitleSum(null);
  this.targetSheet.setColumnWidth(1, 1);
  this.targetSheet.setColumnWidth(2, 500);
//  this.targetSheet.autoResizeColumns(3, this.targetSheet.getMaxColumns());
}

BudgetBuilder.prototype.onTitle = function(row) {
  this.currentTitle = row.getTitle();
  trace("BudgetBuilder.onTitle " + this.currentTitle);
  if (this.targetRowOffset > 0) { // This is not the first title
    if (this.currentTitleSum == 0) { // Section with no content - back up instead of moving on
      --this.targetRowOffset;
    }
    else {
      this.getNextTargetRow(); // Leave one blank row
    }
  }
  var targetRow = this.getNextTargetRow();
  var column = 1;
  targetRow.getCell(1,column++).setValue(this.currentTitle);
  targetRow.getCell(1,column++).setValue(this.currentTitle);
  ++column;
  ++column;
  this.fillTitleSum(targetRow.getCell(1,column++));
  targetRow.setFontWeight("bold");
  targetRow.setFontSize(12);
  targetRow.setBackground("#f3f3f3");
}

BudgetBuilder.prototype.onRow = function(row) {
  var totalPrice = row.getTotalPrice();
  if (Math.abs(totalPrice) > 0) { // This is an item for the invoice
    trace("BudgetBuilder.onRow " + row.getDescription());
    var description = row.getDescription();
    var unitPrice = row.getUnitPrice();
    if (unitPrice == 0.01) { // This is a sub-item, not to be added to the sum
      description = '- ' + description;
      unitPrice = "";
      totalPrice = "";
    }
    else { // This is a top-level item, add to the sum
          this.currentTitleSum += totalPrice;
    }
    var targetRow = this.getNextTargetRow();
    var column = 1;
    targetRow.getCell(1,column++).setValue(this.currentTitle);
    targetRow.getCell(1,column++).setValue(description);
    targetRow.getCell(1,column++).setValue(row.getQuantity());
    targetRow.getCell(1,column++).setValue(unitPrice).setNumberFormat("£#,##0");
    targetRow.getCell(1,column++).setValue(totalPrice).setNumberFormat("£#,##0");
    targetRow.setFontWeight("normal");
    targetRow.setFontSize(10);
    targetRow.setBackground("#ffffff"); // White
    targetRow.set
  } else {
    trace("BudgetBuilder.onRow - ignore (no price): " + row.getDescription());
  }
}

BudgetBuilder.prototype.fillTitleSum = function(nextTitleSumCell) {
  if (this.currentTitleSumCell != null) {
    this.currentTitleSumCell.setValue(this.currentTitleSum).setNumberFormat("£#,##0");
    this.currentTitleSum = 0;
  }
  this.currentTitleSumCell = nextTitleSumCell;
}

// private method getNextTargetRow
//
BudgetBuilder.prototype.getNextTargetRow = function() {
  var targetRow = this.targetRange.offset(this.targetRowOffset++, 0, 1); // A range of 1 row height
  targetRow.breakApart().setFontWeight("normal").setFontSize(10).setBackground("#ffffff");
  if (targetRow.getRowIndex() > targetRow.getSheet().getMaxRows()-2) { // We're at the end, extend
    targetRow.getSheet().insertRowBefore(targetRow.getRowIndex());
  }
  return targetRow;
}

BudgetBuilder.prototype.trace = function() {
  return "{BudgetBuilder " + Range.trace(this.targetRange) + "}";
}


