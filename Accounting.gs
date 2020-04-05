
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
    this.targetRangeName = targetRangeName;
    this.targetRange = CRange.getByName(targetRangeName);
    this.targetSheet = this.targetRange.sheet;
    this.targetRowOffset = 0;
    trace(`NEW ${this.trace}`);
  }

  onBegin() {
    trace("SupplierAccountBuilder.onBegin - reset context");
  }

  onEnd() {
    trace("SupplierAccountBuilder.onEnd - fill final title sum & autofit");
  }

  onTitle(row) {
    this.currentTitle = row.title;
  }
  
  onRow(row) {
    let supplier = row.supplier;
    if (supplier != "") {
      trace(`SupplierAccountBuilder.onRow ${supplier}`);
    }
  }

  get trace() {
    return `{SupplierAccountBuilder ${this.targetRange.trace}}`;
  }
}
