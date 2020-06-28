
function onUpdateSupplierSummary() {
  trace("onUpdateSupplierSummary");
  let eventDetails = new EventDetails();
  let supplierSummaryBuilder = new SupplierSummaryBuilder("SupplierSummary");
  eventDetails.apply(supplierSummaryBuilder);
}

//=============================================================================================
// Class SupplierSummaryBuilder
//
class SupplierSummaryBuilder {
  
  constructor(targetRangeName) {
    this.targetRange = Range.getByName(targetRangeName);
    this.targetSheet = this.targetRange.getSheet();
    this.targetRowOffset = 0;
    this.maxTargetRows = 1000;
    trace("NEW " + this.trace);
  }

  onBegin() {
    trace("SupplierSummaryBuilder.onBegin - reset context");
    this.currentTitle = "";
    this.currentTitleSumCell = null;
    this.currentTitleSum = 0;
    this.targetRange.clear();
    let column3 = this.targetRange.offset(0, 2, this.targetRange.getNumRows(), 1); // Select a full column
    let column4 = column3.offset(0, 1);
    let column5 = column4.offset(0, 1);
    column3.setNumberFormat("0");
    column4.setNumberFormat("£#,##0.00");
    column5.setNumberFormat("£#,##0.00");
    this.targetRange.setFontWeight("normal").setFontSize(10);
  }

  onEnd() {
    trace("SupplierSummaryBuilder.onEnd - fill final title sum & autofit");
    this.fillTitleSum(null);
    this.targetSheet.setColumnWidth(1, 1);
    this.targetSheet.setColumnWidth(2, 500);
    //  this.targetSheet.autoResizeColumns(3, this.targetSheet.getMaxColumns());
  }

  onTitle(row) {
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

  onRow(row) {
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
  
  fillTitleSum(nextTitleSumCell) {
    if (this.currentTitleSumCell != null) {
      this.currentTitleSumCell.setValue(this.currentTitleSum);
      this.currentTitleSum = 0;
    }
    this.currentTitleSumCell = nextTitleSumCell;
  }

  // private method getNextTargetRow
  //
  getNextTargetRow() {
    return this.targetRange.offset(Math.min(this.targetRowOffset++, this.maxTargetRows-1), 0, 1); // A range of 1 row height
  }

  get trace() {
    return "{SupplierSummaryBuilder " + Range.trace(this.targetRange) + "}";
  }
}
