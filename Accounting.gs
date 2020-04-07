
function onUpdateSupplierAccountSummary() {
  trace("onUpdateSupplierAccountSummary");
  let eventDetailsIterator = new EventDetailsIterator();
  let supplierAccountBuilder = new SupplierAccountBuilder("SupplierAccountSummary");
  eventDetailsIterator.iterate(supplierAccountBuilder);
  
}

//=============================================================================================
// Class SupplierAccountBuilder
//
class SupplierAccountBuilder {
  constructor(targetRangeName) {
    this.targetRange = CRange.getByName(targetRangeName);
    this.targetSheet = this.targetRange.sheet;
    this.targetRowOffset = 0;
    trace(`NEW ${this.trace}`);
  }

  onBegin() {
    trace("SupplierAccountBuilder.onBegin - reset context");
    this.suppliers = new SupplierList;
  }

  onEnd() {
    trace("SupplierAccountBuilder.onEnd - fill final title sum & autofit");
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
    let targetRow = this.targetRange.range.offset(this.targetRowOffset++, 0, 1); // A range of 1 row height
    targetRow.breakApart().setFontWeight("normal").setFontSize(10).setBackground("#ffffff");
    targetRow.getSheet().insertRowAfter(targetRow.getRowIndex());
    return targetRow;
  }

  get trace() {
    return `{SupplierAccountBuilder ${this.targetRange.trace}}`;
  }
}


//=============================================================================================
// Class SupplierList

class SupplierList {
  constructor() {
    this.myList = [];
  }
  
  add(supplier) {
    //if this.myList[supplier];
    this.myList.push(supplier);
  }
  
  sortUnique() {
    this.myList = [...new Set(this.myList)]; 
    this.myList.sort();
  }

  get list() { return this.myList; }
  
  get trace() {
    let text = "";
    this.myList.forEach((item) => { text = text + item + " " })
    return text;
  }
}
