function onUpdateSupplierCosting() {
  trace("onUpdateSupplierCosting");
  let eventDetails = new EventDetails();
  let supplierCostingBuilder = new SupplierCostingBuilder(Range.getByName("SupplierCosting", "Supplier Costing"));
  eventDetails.sort(SortType.supplier);
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
    this.currentSupplierSumCell = null;
    this.currentSupplierSum = 0;
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
    if (this.currentSupplierSum == 0) {     // Last section is empty?
      this.targetRange.getPreviousRow(); //  yes - Back up one row (last title will be trimmed away)
    }
    else {
      this.fillSupplierSum(null);
    }
//  this.targetSheet.nativeSheet.setColumnWidth(1, 100);
//  this.targetSheet.nativeSheet.setColumnWidth(2, 500);
//  this.targetSheet.autoResizeColumns(3, this.targetSheet.getMaxColumns());
    this.targetRange.trim(); // delete excessive rows at the end
  }

  onTitle(row) {
    trace("SupplierCostingBuilder.onTitle " + this.currentTitle + " - ignore");
  }
  
  newSupplierSection(row) {
    this.currentSupplier = row.supplier;
    ++this.currentSection;
    trace("SupplierCostingBuilder.newSupplierSection " + this.suppler);
    if (this.currentSection > 1) { // This is not the first section (if it is, no need for clean-up house-keeping)
      if (this.currentSupplierSum == 0) {       // Section with no content?
        this.targetRange.getPreviousRow();      //  - back up instead of moving on
      }
      else {
        this.targetRange.getNextRowAndExtend(); //  - else leave one blank row before next section
      }
    }
    let targetRow = this.targetRange.getNextRowAndExtend();
    let column = 1;
    targetRow.getCell(1,column++).setValue(this.currentSupplier);
    targetRow.getCell(1,column++).setValue(this.currentSupplier);
    ++column;
    ++column;
    this.fillSupplierSum(targetRow.getCell(1,column++));
    targetRow.setFontWeight("bold");
    targetRow.setFontSize(12);
    targetRow.setBackground("#f3f3f3");
  }

  onRow(row) {
    var totalCost = row.totalCost;
    if (Math.abs(totalCost) > 0) { // This is an item for the invoice
      trace("SupplierCostingBuilder.onRow " + row.description);
      if (row.supplier !== this.currentSupplier) {
        this.newSupplierSection(row);
      }
      let description = row.description;
      let unitCost = row.unitCost;
      if (row.isSubItem) { // This is a sub-item, not to be added to the sum
        description = "- " + description;
        unitCost = "";
        totalCost = "";
      }
      else { // This is a top-level item, add to the sum
        this.currentSupplierSum += totalCost;
      }
      let targetRow = this.targetRange.getNextRowAndExtend();
      let column = 1;
      targetRow.getCell(1,column++).setValue(row.supplier);
//      targetRow.getCell(1,column++).setValue(this.currentTitle);
      targetRow.getCell(1,column++).setValue(description);
      targetRow.getCell(1,column++).setValue(row.quantity);
      targetRow.getCell(1,column++).setValue(unitCost).setNumberFormat("£#,##0.00");
      targetRow.getCell(1,column++).setValue(totalCost).setNumberFormat("£#,##0.00");
      targetRow.setFontWeight("normal");
      targetRow.setFontSize(10);
      targetRow.setBackground("#ffffff"); // White
    } else {
      trace("SupplierCostingBuilder.onRow - ignore (no price): " + row.description);
    }
  }

  fillSupplierSum(nextSupplierSumCell) {
    if (this.currentSupplierSumCell != null) {
      this.currentSupplierSumCell.setValue(this.currentSupplierSum).setNumberFormat("£#,##0.00");
      this.currentSupplierSum = 0;
    }
    this.currentSupplierSumCell = nextSupplierSumCell;
  }

  get trace() {
    return "{SupplierCostingBuilder " + this.targetRange.trace + "}";
  }
}
