function onUpdateAccountSummary() {
  trace("onUpdateAccountSummary");
  let eventDetailsIterator = new EventDetailsIterator();
  let accountSummaryBuilder = new AccountSummaryBuilder(Range.getByName("SupplierAccountSummary", "Accounting"));
  eventDetailsIterator.iterate(accountSummaryBuilder);
  
}

//=============================================================================================
// Class SupplierAccountBuilder
//
class AccountSummaryBuilder {
  constructor(targetRange) {
    this.targetRange = targetRange;
    this.targetSheet = this.targetRange.sheet;
    this.targetRowOffset = 0;
    trace(`NEW ${this.trace}`);
  }

  onBegin() {
    trace("AccountSummaryBuilder.onBegin - reset context");
    this.suppliers = new SupplierList;
  }

  onEnd() {
    trace("AccountSummaryBuilder.onEnd - fill final title sum & autofit");
    this.suppliers.sortUnique();
    trace("Suppliers: " + this.suppliers.trace);
    this.targetRange.deleteExcessiveRows(2); // Keep 2 rows
    this.targetRange.clear();
    for (var supplier of this.suppliers.list) {
      let targetRow = this.getNextTargetRow();
      let targetRowIndex = targetRow.getRow();
      let column = 1;
      targetRow.getCell(1,column++).setValue(supplier);
      targetRow.getCell(1,column++).setValue(`=VLOOKUP(A${targetRowIndex},Coordination!L15:O160,4,FALSE)`);        
      targetRow.getCell(1,column++).setValue(`=sumif(Coordination!L15:L160, A${targetRowIndex}, Coordination!X15:X160)`);    
    } 
  }

  onTitle(row) {
    this.currentTitle = row.title;
  }
  
  onRow(row) {
    let supplier = row.supplier;
    if (supplier != "") {
      trace(`SupplierAccountBuilder.onRow ${supplier}`);
      this.suppliers.add(supplier);
    }
  }

  getNextTargetRow() {
    let targetRow = this.targetRange.nativeRange.offset(this.targetRowOffset++, 0, 1); // A range of 1 row height
    targetRow.breakApart().setFontWeight("normal").setFontSize(10).setBackground("#ffffff");
    targetRow.getSheet().insertRowAfter(targetRow.getRowIndex());
    return targetRow;
  }

  get trace() {
    return `{AccountSummaryBuilder ${this.targetRange.trace}}`;
  }
}


//=============================================================================================
// Class SupplierList

class SupplierList {
  constructor() {
    this._list = [];
  }
  
  add(supplier) {
    //if this._list[supplier];
    this._list.push(supplier);
  }
  
  sortUnique() {
    this._list = [...new Set(this._list)]; 
    this._list.sort();
  }

  get list() { return this._list; }
  
  get trace() {
    let text = "";
    this._list.forEach((item) => { text = text + item + " " })
    return text;
  }
}
