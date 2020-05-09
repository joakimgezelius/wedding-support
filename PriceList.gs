function onRefreshPriceList() {
  trace("onRefreshPriceList");
  //let priceListIterator = new EventDetailsIterator();
//  let eventDetailsUpdater = new EventDetailsUpdater(false);
}

function onUpdatePackages() {
  trace("onUpdatePackages");
//  if (Dialog.confirm("Forced Coordinator Update - Confirmation Required", "Are you sure you want to force-update the coordinator? It will overwrite row numbers and formulas, make sure the sheet is sorted properly!") == true) {
//    let eventDetailsIterator = new EventDetailsIterator();
//    let eventDetailsUpdater = new EventDetailsUpdater(true);
//    eventDetailsIterator.iterate(eventDetailsUpdater);
//  }
}

function onPriceListClearExport() {
  trace("onPriceListClearExport");
  let priceListExport = new PriceListExport("Export");
  priceListExport.clear();
}

function onPriceListClearSelectionTicks() {
  trace("onPriceListClearSelectionTicks");
  let priceList = new PriceList("PriceList");
  if (Dialog.confirm("Clear Selection Ticks", 'Are you sure you want to clear all ticks in the "Selected for Quote" column?') == true) {
    priceList.clearSelectionTicks();
  }
}

function onPriceListExportTicked() {
  trace("onPriceListExport");
  let priceList = new PriceList("PriceList");
  let priceListExport = new PriceListExport("Export");
  priceList.iterate(priceListExport);
}

function onPriceListExportSelection() {
  trace("onPriceListExport");
  let priceList = new PriceList("PriceList");
  let priceListExport = new PriceListExport("Export");
  priceList.iterate(priceListExport);
}


//================================================================================================

class PriceList {
  
  constructor(rangeName) {
    this.range = Range.getByName(rangeName).loadColumnNames();
    this.rowCount = this.range.height;
    this.values = this.range.values;
    this.formulas = this.range.formulas;
    trace("NEW " + this.trace);
  }
  
  update() {
    this.gatherCategories();
  }

  // Method iterate
  // Iterate over all price list rows
  //
  iterate(handler) {
    trace("PriceList.iterate " + this.trace);
    handler.onBegin();
    for (let rowOffset = 0; rowOffset < this.rowCount; rowOffset++) {
      let row = new PriceListRow(this.values[rowOffset], this.formulas[rowOffset], rowOffset, this.range);
      if (row.isTitle) {
        handler.onTitle(row);
      } else {
        handler.onRow(row);
      }
    }  
    handler.onEnd();
  }

  clearSelectionTicks() {
    trace("PriceList.clearSelectionTicks " + this.trace);
    for (let rowOffset = 0; rowOffset < this.rowCount; rowOffset++) {
      let row = new PriceListRow(this.values[rowOffset], this.formulas[rowOffset], rowOffset, this.range);
      if (row.isSelected) {
        row.isSelected = false;
      }
    }  
  }
  
  gatherCategories() {
    trace("PriceList.gatherCategories");
    for (var rowOffset = 0; rowOffset < this.rowCount; rowOffset++) {
      var rowRange = this.range.offset(rowOffset, 0, 1);
    }
  }

  get trace() {
    return "{PriceList " + this.range.trace + "}";
  }

} // PriceList


//================================================================================================

class PriceListRow extends EventRow {

  constructor(values, formulas, rowOffset, containerRange) {
    super(values, formulas, rowOffset, containerRange);
  }

  get isSelected()      { return this.get("Selected"); } // Accept blank
  set isSelected(value) { this.set("Selected", value); }
  
} // PriceListRow


//================================================================================================

class PriceListExport {
  
  constructor(rangeName) {
    this.range = Range.getByName(rangeName).loadColumnNames();
    this.values = this.range.values;
    this.formulas = this.range.formulas;
    trace("NEW " + this.trace);
  }
  
  clear() {
    trace("PriceListExport.clear " + this.trace);
    this.range.minimizeAndClear();
  }
  
  onBegin() {
    trace("PriceListExport.onBegin " + this.trace);
  }
  
  onEnd() {
    trace("PriceListExport.onEnd " + this.trace);
  }
  
  onTitle(row) {
    trace("PriceListExport.onTitle - ignore");
  }
  
  onRow(priceListRow) {
    trace("PriceListExport.onRow ");
    if (priceListRow.isSelected) {
      let exportRowRange = this.range.getNextRowAndExtend();
      let rowOffset = this.range.currentRowOffset
      let exportRow = new EventRow(this.values[rowOffset], this.formulas[rowOffset], rowOffset, this.range);
      exportRow.category             = priceListRow.category;
      exportRow.supplier             = priceListRow.supplier;
      exportRow.description          = priceListRow.description;
      exportRow.currency             = priceListRow.currency;
      exportRow.nativeUnitCost       = priceListRow.nativeUnitCost;
      exportRow.markup               = priceListRow.markup;
      exportRow.commissionPercentage = priceListRow.commissionPercentage;

    }    
  }
  
  get trace() {
    return "{PriceListExport " + this.range.trace + "}";
  }

} // PriceListExport


