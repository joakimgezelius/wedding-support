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
    this.currentSectionRows = 0;
    this.currentSupplier = "(none)";
    this.currentSupplierCurrency = "";
    this.currentSupplierCurrencyFormat = Helper.currencyFormat(this.currentSupplierCurrency);
    this.currentSupplierGrossSum = 0;
    this.currentSupplierCommissionSum = 0;
    this.currentSupplierClientPriceSum = 0;
    this.currentSupplierMarginSum = 0;
    this.eurGbpRate = Spreadsheet.getCellValue("EURGBP");
    trace("SupplierCostingBuilder.onBegin - EURGBP=" + this.eurGbpRate);
    this.targetRange.minimizeAndClear(SupplierCostingBuilder.formatRange); // Keep 2 rows
    //let column3 = this.targetRange.nativeRange.offset(0, 2, this.targetRange.height, 1); // Select a full column
    //let column4 = column3.offset(0, 1);
    //let column5 = column4.offset(0, 1);
    //column3.setNumberFormat("0");
    //column4.setNumberFormat("£#,##0.00");
    //column5.setNumberFormat("£#,##0.00");
    trace("SupplierCostingBuilder.onBegin - done");
  }

  onEnd() {
    trace("SupplierCostingBuilder.onEnd - fill final supplier sum, autofit & trim");
    if (this.currentSectionRows == 0) {   // Last section is empty?
      this.targetRange.getPreviousRow();  //  yes - Back up one row (last title will be trimmed away)
    }
    else {
      this.fillSectionHeader();
    }
//  this.targetSheet.nativeSheet.setColumnWidth(1, 100);
//  this.targetSheet.nativeSheet.setColumnWidth(2, 500);
//  this.targetSheet.autoResizeColumns(3, this.targetSheet.getMaxColumns());
    this.targetRange.trim(); // delete excessive rows at the end
  }

  onTitle(row) {
    trace("SupplierCostingBuilder.onTitle " + row.title + " - ignore");
  }
  
   onRow(row) {
    if (this.isRowToBeIgnored(row)) {  // Check if this row is to be ignored
      trace("SupplierCostingBuilder.onRow - ignore: " + row.description);
    } 
    else {
      if (row.supplier !== this.currentSupplier) { // Check if we've hit the next supplier, if so, insert header etc.
        this.newSupplierSection(row);
      }
      this.currentSectionRows++;
      trace(`SupplierCostingBuilder.onRow ${this.currentSection}:${this.currentSectionRows} ${row.description}`);

      var totalNativeGrossCost = Number(row.totalNativeGrossCost);
      this.currentSupplierGrossSum += totalNativeGrossCost; // Summing gross cost for the supplier sub-header

      var totalClientPrice = Number(row.totalPrice);
      this.currentSupplierClientPriceSum += Number(totalClientPrice);

      // Commission calculations
      var commissionPercentage = Number(row.commissionPercentage);
      var commissionAmount = commissionPercentage * totalNativeGrossCost;
      this.currentSupplierCommissionSum += commissionAmount;  // Summing commission for the supplier sub-header

      // Margin calculations
      var margin = totalClientPrice - totalNativeGrossCost / (row.currency == "GBP" ? 1 : this.eurGbpRate) + commissionAmount;
      this.currentSupplierMarginSum += margin;  // Summing margin for the supplier sub-header
      var marginPercentage = (totalClientPrice != 0 ? margin / totalClientPrice : "-");

      let targetRow = this.targetRange.getNextRowAndExtend();
      let column = 1;
      targetRow.getCell(1,column++).setValue(row.description);
      targetRow.getCell(1,column++).setValue(row.status);
      targetRow.getCell(1,column++).setValue(row.quantity);  // Quantity
      targetRow.getCell(1,column++).setValue(row.nativeUnitCostWithVAT).setNumberFormat(row.currencyFormat);
      targetRow.getCell(1,column++).setValue(totalNativeGrossCost).setNumberFormat(row.currencyFormat);
      targetRow.getCell(1,column++).setValue(totalClientPrice).setNumberFormat("£ #,##0.00");
      
      if (commissionPercentage !== 0) { // Only show if there is a commission
        targetRow.getCell(1,column++).setValue(commissionPercentage);  // Commission %
        targetRow.getCell(1,column++).setValue(commissionAmount).setNumberFormat(row.currencyFormat);  // Commission amount
      } 
      else {
        ++column;
        ++column;
      }
      targetRow.getCell(1,column++).setValue(marginPercentage);
      targetRow.getCell(1,column++).setValue(margin);
      targetRow.getCell(1,column++).setValue(row.itemNo);
      targetRow.getCell(1,column++).setValue(row.currency);  // Currency
      targetRow.getCell(1,column++).setValue(row.supplier);
      targetRow.getCell(1,column++).setValue(row.location);
      targetRow.setFontWeight("normal");
      targetRow.setFontSize(10);
      targetRow.setBackground("#ffffff"); // White
    }
  }

  isRowToBeIgnored(row) { 
    if (row.supplier === "") return true;         // Ignore rows with no supplier 
    if (row.status == "Cancelled") return true;   // Ignore cancelled rows 
//  if (!row.isSelected) return true; // Ignore rows that are not selected
//  if (row.isSubItem) return true;
//  if (Math.abs(row.nativeUnitCostWithVAT) > 0.01) return true;
//  if (row.quantity == 0) return true;
    return false;
  }

  newSupplierSection(row) {
    ++this.currentSection;
    trace(`SupplierCostingBuilder.newSupplierSection ${this.currentSection}: ${row.supplier}`);
    if (this.currentSection == 1) { // This is the first section, no need for back-filling
      trace(`SupplierCostingBuilder.newSupplierSection first section, no filling`);
    } 
    else {
      if (this.currentSectionRows == 0) {       // Section with no content?
        this.targetRange.getPreviousRow();      //  - back up instead of moving on
      }
      else {
        this.targetRange.getNextRowAndExtend(); //  - else leave one blank row before next section
      }
      this.fillSectionHeader();
    }
    // Initialise context for the new section 
    this.sectionHeaderRow = this.targetRange.getNextRowAndExtend();
    this.currentSupplier = row.supplier;
    this.currentSupplierCurrency = row.currency;
    this.currentSupplierCurrencyFormat = Helper.currencyFormat(this.currentSupplierCurrency);
    this.currentSectionRows = 0;
    this.currentSupplierGrossSum = 0;
    this.currentSupplierCommissionSum = 0;
    this.currentSupplierClientPriceSum = 0;
    this.currentSupplierMarginSum = 0;
  }

  fillSectionHeader() {
    let headerRow = this.sectionHeaderRow;
    headerRow.getCell(1,13).setValue(this.currentSupplier);  // Set supplier in supplier column
    headerRow.getCell(1,1).setValue(this.currentSupplier);   // Also set supplier header in column 1 (Description)
    headerRow.getCell(1,5).setValue(this.currentSupplierGrossSum).setNumberFormat(this.currentSupplierCurrencyFormat);
    if (this.currentSupplierCommissionSum > 0) { // Only fill in a commission sum if non-zero
      headerRow.getCell(1,8).setValue(this.currentSupplierCommissionSum).setNumberFormat(this.currentSupplierCurrencyFormat);
    }
    headerRow.getCell(1,6).setValue(this.currentSupplierClientPriceSum).setNumberFormat("£ #,##0.00");
    headerRow.getCell(1,10).setValue(this.currentSupplierMarginSum).setNumberFormat("£ #,##0.00");
    let marginPercentage = (this.currentSupplierClientPriceSum != 0 ? this.currentSupplierMarginSum / this.currentSupplierClientPriceSum : "-");
    headerRow.getCell(1,9).setValue(marginPercentage).setNumberFormat("0%");

    // Format the row
    headerRow.setFontWeight("bold");
    headerRow.setFontSize(12);
    headerRow.setBackground("#f3f3f3");
  }

  get trace() {
    return `{SupplierCostingBuilder ${this.targetRange.trace}}`;
  }
}
