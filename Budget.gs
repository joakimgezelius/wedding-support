function onUpdateBudget() {
  trace("onUpdateBudget");
  let eventDetails = new EventDetails();
  let budgetBuilder = new BudgetBuilder(Range.getByName("Budget", "Budget"));
  eventDetails.apply(budgetBuilder);
}

//=============================================================================================
// Class BudgetBuilder
//
class BudgetBuilder {
  constructor(targetRange) {
    this.targetRange = targetRange;
    this.targetSheet = this.targetRange.sheet;
    trace("NEW " + this.trace);
  }

  static formatRange(range) {
    trace("BudgetBuilder.formatRange");
    range.setFontWeight("normal")
    .setFontSize(10)
    .breakApart()
    .setBackground("#ffffff")
    .setWrap(true);
  };
  
  onBegin() {
    trace("BudgetBuilder.onBegin - reset context");
    this.currentSection = 0;
    this.currentTitle = "";
    this.currentTitleSumCell = null;
    this.currentTitleSum = 0;
    this.targetRange.minimizeAndClear(BudgetBuilder.formatRange); // Keep 2 rows
    let column3 = this.targetRange.nativeRange.offset(0, 2, this.targetRange.height, 1); // Select a full column
    let column4 = column3.offset(0, 1);
    let column5 = column4.offset(0, 1);
    column3.setNumberFormat("0");
    column4.setNumberFormat("£#,##0.00");
    column5.setNumberFormat("£#,##0.00");
    trace("BudgetBuilder.onBegin - done");
  }

  onEnd() {
    trace("BudgetBuilder.onEnd - fill final title sum, autofit & trim");
    if (this.currentTitleSum == 0) {     // Last section is empty?
      this.targetRange.getPreviousRow(); //  yes - Back up one row (last title will be trimmed away)
    }
    else {
      this.fillTitleSum(null);
    }
    this.targetSheet.nativeSheet.setColumnWidth(1, 1);
    this.targetSheet.nativeSheet.setColumnWidth(2, 500);
//  this.targetSheet.autoResizeColumns(3, this.targetSheet.getMaxColumns());
    this.targetRange.trim(); // delete excessive rows at the end
  }

  onTitle(row) {
    this.currentTitle = row.title;
    ++this.currentSection;
    trace("BudgetBuilder.onTitle " + this.currentTitle);
    if (this.currentSection > 1) { // This is not the first section (if it is, no need for clean-up house-keeping)
      if (this.currentTitleSum == 0) {          // Section with no content?
        this.targetRange.getPreviousRow();      //  - back up instead of moving on
      }
      else {
        this.targetRange.getNextRowAndExtend(); //  - else leave one blank row before next section
      }
    }
    let targetRow = this.targetRange.getNextRowAndExtend();
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
    const totalPrice = row.totalPrice;
    const category = row.category; 
    if (Math.abs(totalPrice) > 0 && 
        category != "Labour" && 
        category != "Transport" &&
        category != "Styling"
    ) { // This is an item for the invoice
      trace("BudgetBuilder.onRow " + row.description);
      let description = row.description;
      let unitPrice = row.unitPrice;
      if (row.isSubItem) { // This is a sub-item, not to be added to the sum
        description = "- " + description;
        unitPrice = "";
        totalPrice = "";
      }
      else { // This is a top-level item, add to the sum
        this.currentTitleSum += totalPrice;
      }
      let targetRow = this.targetRange.getNextRowAndExtend();
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

  get trace() {
    return "{BudgetBuilder " + this.targetRange.trace + "}";
  }
}
