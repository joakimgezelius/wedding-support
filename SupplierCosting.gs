function onUpdateSupplierCosting() {
  trace("onUpdateSupplierCosting");
  let eventDetails = new EventDetails();
  let supplierCostingBuilder = new SupplierCostingBuilder(Range.getByName("SupplierCosting", "Supplier Costing"));
  eventDetails.sort(SortType.supplier_location);
  eventDetails.apply(supplierCostingBuilder);
}

//=============================================================================================
// Class SupplierCostingBuilder
//
class SupplierCostingBuilder {
  constructor(targetRange) {
    this.targetRange = targetRange;
    this.targetSheet = this.targetRange.sheet;
    trace("NEW " + this.trace);
  }

  static formatRange(range) {
    trace("SupplierCostingBuilder.formatRange");
    range.setFontWeight("normal")
    .setFontSize(10)
    .breakApart()
    .setBackground("#ffffff")
    .setWrap(true);
  };
  
  onBegin() {
    trace("SupplierCostingBuilder.onBegin - reset context");
    this.currentSection = 0;
    this.currentSupplier = "(none)";
    this.currentSupplierGrossEurSumCell = null;
    this.currentSupplierGrossEurSum = 0;
    this.currentSupplierGrossGbpSumCell = null;
    this.currentSupplierGrossGbpSum = 0;
    this.currentSupplierNettSumCell = null;
    this.currentSupplierNettSum = 0;
    this.targetRange.minimizeAndClear(SupplierCostingBuilder.formatRange); // Keep 2 rows
    let column3 = this.targetRange.nativeRange.offset(0, 2, this.targetRange.height, 1); // Select a full column
    let column4 = column3.offset(0, 1);
    let column5 = column4.offset(0, 1);
    column3.setNumberFormat("0");
    column4.setNumberFormat("£#,##0.00");
    column5.setNumberFormat("£#,##0.00");
    trace("SupplierCostingBuilder.onBegin - done");
  }

  onEnd() {
    trace("SupplierCostingBuilder.onEnd - fill final supplier sum, autofit & trim");
    if (this.currentSupplierNettSum == 0) {     // Last section is empty?
      this.targetRange.getPreviousRow(); //  yes - Back up one row (last title will be trimmed away)
    }
    else {
      this.fillSupplierSums(null, null, null);
    }
//  this.targetSheet.nativeSheet.setColumnWidth(1, 100);
//  this.targetSheet.nativeSheet.setColumnWidth(2, 500);
//  this.targetSheet.autoResizeColumns(3, this.targetSheet.getMaxColumns());
    this.targetRange.trim(); // delete excessive rows at the end
  }

  onTitle(row) {
    this.currentTitle = row.title;
    trace("SupplierCostingBuilder.onTitle " + this.currentTitle + " - ignore");
  }
  
  newSupplierSection(row) {
    this.currentSupplier = row.supplier;
    ++this.currentSection;
    trace("SupplierCostingBuilder.newSupplierSection " + this.suppler);
    if (this.currentSection > 1) { // This is not the first section (if it is, no need for clean-up house-keeping)
      if (this.currentSupplierGrossSum == 0) {       // Section with no content?
        this.targetRange.getPreviousRow();      //  - back up instead of moving on
      }
      else {
        this.targetRange.getNextRowAndExtend(); //  - else leave one blank row before next section
      }
    }
    let targetRow = this.targetRange.getNextRowAndExtend();
    let column = 1;
    ++column;
    targetRow.getCell(1,column++).setValue(this.currentSupplier);
    targetRow.getCell(1,column++).setValue(this.currentSupplier);
    ++column;
    ++column;
    ++column;
    ++column;
    this.fillSupplierSums(targetRow.getCell(1,column++), targetRow.getCell(1,column++), targetRow.getCell(1,column++));
    targetRow.setFontWeight("bold");
    targetRow.setFontSize(12);
    targetRow.setBackground("#f3f3f3");
  }

  onRow(row) {
    var totalNativeGrossCost = row.totalNativeGrossCost;
    var totalNettCost = row.totalNettCost;
    var paymentMethod = row.paymentMethod;
    var paymentStatus = row.paymentStatus;
    if (!row.isSubItem && Math.abs(row.nativeUnitCostWithVAT) > 0.01 && row.quantity > 0) { // This not a sub-item (marked as such or with no price)
      trace("SupplierCostingBuilder.onRow " + row.description);
      if (row.supplier !== this.currentSupplier) {
        this.newSupplierSection(row);
      }
      let description = row.description;
      this.currentSupplierNettSum += Number(totalNettCost);
      let targetRow = this.targetRange.getNextRowAndExtend();
      let column = 1;
      targetRow.getCell(1,column++).setValue(row.itemNo);
      targetRow.getCell(1,column++).setValue(row.supplier);
//    targetRow.getCell(1,column++).setValue(this.currentTitle);
      targetRow.getCell(1,column++).setValue(row.status);
      targetRow.getCell(1,column++).setValue(row.location);
      targetRow.getCell(1,column++).setValue(description);
      targetRow.getCell(1,column++).setValue(row.quantity);
      targetRow.getCell(1,column++).setValue(row.nativeUnitCostWithVAT).setNumberFormat(row.currencyFormat);
      if (row.currency === "GBP") {
        this.currentSupplierGrossGbpSum += Number(totalNativeGrossCost);
        ++column; // Skip EUR column
        targetRow.getCell(1,column++).setValue(totalNativeGrossCost).setNumberFormat("£#,##0.00");
      } else {
        this.currentSupplierGrossEurSum += Number(totalNativeGrossCost);
        targetRow.getCell(1,column++).setValue(totalNativeGrossCost).setNumberFormat("€#,##0.00");
        ++column; // Skip GBP column
      }
      targetRow.getCell(1,column++).setValue(totalNettCost).setNumberFormat("£#,##0.00");
      targetRow.getCell(1,column++).setValue(paymentMethod);
      targetRow.getCell(1,column++).setValue(paymentStatus);
      targetRow.setFontWeight("normal");
      targetRow.setFontSize(10);
      targetRow.setBackground("#ffffff"); // White
    } else {
      trace("SupplierCostingBuilder.onRow - ignore (no price): " + row.description);
    }
  }

  fillSupplierSums(nextSupplierGrossEurSumCell, nextSupplierGrossGbpSumCell, nextSupplierNettSumCell) {
    if (this.currentSupplierGrossEurSumCell != null) { // This is not the first section
      if (this.currentSupplierGrossEurSum != 0) {
        this.currentSupplierGrossEurSumCell.setValue(this.currentSupplierGrossEurSum).setNumberFormat("€#,##0.00");
        this.currentSupplierGrossEurSum = 0;
      }
      if (this.currentSupplierGrossGbpSum != 0) {
        this.currentSupplierGrossGbpSumCell.setValue(this.currentSupplierGrossGbpSum).setNumberFormat("£#,##0.00");
        this.currentSupplierGrossGbpSum = 0;
      }
      this.currentSupplierNettSumCell.setValue(this.currentSupplierNettSum).setNumberFormat("£#,##0.00");
      this.currentSupplierNettSum = 0;
    }
    this.currentSupplierGrossEurSumCell = nextSupplierGrossEurSumCell;
    this.currentSupplierGrossGbpSumCell = nextSupplierGrossGbpSumCell;
    this.currentSupplierNettSumCell = nextSupplierNettSumCell;
  }

  get trace() {
    return "{SupplierCostingBuilder " + this.targetRange.trace + "}";
  }
}
