
function onUpdateBudget() {
  trace("onUpdateBudget");
  let eventDetailsIterator = new EventDetailsIterator();
  let budgetBuilder = new BudgetBuilder("Budget");
  eventDetailsIterator.iterate(budgetBuilder);
}

//=============================================================================================
// Class BudgetBuilder
//
class BudgetBuilder {
  constructor(targetRangeName) {
    this.targetRange = Range.getByName(targetRangeName);
    this.targetSheet = this.targetRange.sheet;
    this.targetRowOffset = 0;
    trace("NEW " + this.trace);
  }

  onBegin() {
    trace("BudgetBuilder.onBegin - reset context");
    this.currentTitle = "";
    this.currentTitleSumCell = null;
    this.currentTitleSum = 0;
    this.targetRange.deleteExcessiveRows(2); // Keep 2 rows
    this.targetRange.clear();
    this.targetRange.range.setFontWeight("normal").setFontSize(10).breakApart().setBackground("#ffffff").setWrap(true);
    trace("BudgetBuilder.onBegin - format columns");
    let column3 = this.targetRange.range.offset(0, 2, this.targetRange.height, 1); // Select a full column
    let column4 = column3.offset(0, 1);
    let column5 = column4.offset(0, 1);
    column3.setNumberFormat("0");
    column4.setNumberFormat("£#,##0.00");
    column5.setNumberFormat("£#,##0.00");
    trace("BudgetBuilder.onBegin - done");
  }

  onEnd() {
    trace("BudgetBuilder.onEnd - fill final title sum & autofit");
    this.fillTitleSum(null);
    this.targetSheet.setColumnWidth(1, 1);
    this.targetSheet.setColumnWidth(2, 500);
//  this.targetSheet.autoResizeColumns(3, this.targetSheet.getMaxColumns());
  }

  onTitle(row) {
    this.currentTitle = row.title;
    trace("BudgetBuilder.onTitle " + this.currentTitle);
    if (this.targetRowOffset > 0) { // This is not the first title
      if (this.currentTitleSum == 0) { // Section with no content - back up instead of moving on
        --this.targetRowOffset;
      }
      else {
        this.getNextTargetRow(); // Leave one blank row
      }
    }
    let targetRow = this.getNextTargetRow();
    let column = 1;
    targetRow.getCell(1,column++).setValue(this.currentTitle);
    targetRow.getCell(1,column++).setValue(this.currentTitle);
    ++column;
    ++column;
    this.fillTitleSum(targetRow.getCell(1,column++));
    targetRow.setFontWeight("bold");
    targetRow.setFontSize(12);
    targetRow.setBackground("#f3f3f3");
  }

  onRow(row) {
    var totalPrice = row.totalPrice;
    if (Math.abs(totalPrice) > 0) { // This is an item for the invoice
      trace("BudgetBuilder.onRow " + row.description);
      let description = row.description;
      let unitPrice = row.unitPrice;
      if (unitPrice == 0.01) { // This is a sub-item, not to be added to the sum
        description = "- " + description;
        unitPrice = "";
        totalPrice = "";
      }
      else { // This is a top-level item, add to the sum
        this.currentTitleSum += totalPrice;
      }
      let targetRow = this.getNextTargetRow();
      let column = 1;
      targetRow.getCell(1,column++).setValue(this.currentTitle);
      targetRow.getCell(1,column++).setValue(description);
      targetRow.getCell(1,column++).setValue(row.quantity);
      targetRow.getCell(1,column++).setValue(unitPrice).setNumberFormat("£#,##0");
      targetRow.getCell(1,column++).setValue(totalPrice).setNumberFormat("£#,##0");
      targetRow.setFontWeight("normal");
      targetRow.setFontSize(10);
      targetRow.setBackground("#ffffff"); // White
    } else {
      trace("BudgetBuilder.onRow - ignore (no price): " + row.description);
    }
  }

  fillTitleSum(nextTitleSumCell) {
    if (this.currentTitleSumCell != null) {
      this.currentTitleSumCell.setValue(this.currentTitleSum).setNumberFormat("£#,##0");
      this.currentTitleSum = 0;
    }
    this.currentTitleSumCell = nextTitleSumCell;
  }

// private method getNextTargetRow
//
  getNextTargetRow() {
    let targetRow = this.targetRange.range.offset(this.targetRowOffset++, 0, 1); // A range of 1 row height
    targetRow.breakApart().setFontWeight("normal").setFontSize(10).setBackground("#ffffff");
    if (targetRow.getRowIndex() > targetRow.getSheet().getMaxRows()-2) { // We're at the end, extend
      targetRow.getSheet().insertRowBefore(targetRow.getRowIndex());
    }
    return targetRow;
  }

  get trace() {
    return "{BudgetBuilder " + this.targetRange.trace + "}";
  }
}
