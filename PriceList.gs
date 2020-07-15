PriceListSpreadsheetId = "1lunFhyOgQL1au5JmwWoLSuebPgxLwXqymzGc5FJJ-IU";
PriceListRangeName = "PriceList";

function onRefreshPriceList() {
  trace("onRefreshPriceList");
}

function onUpdatePackages() {
  trace("onUpdatePackages");
//  if (Dialog.confirm("Forced Coordinator Update - Confirmation Required", "Are you sure you want to force-update the coordinator? It will overwrite row numbers and formulas, make sure the sheet is sorted properly!") == true) {
//  }
}

function onPriceListClearExport() {
  trace("onPriceListClearExport");
  let priceListExport = new PriceListExport("Export");
  priceListExport.clear();
}

function onPriceListClearSelectionTicks() {
  trace("onPriceListClearSelectionTicks");
  let priceList = new PriceList();
  if (Dialog.confirm("Clear Selection Ticks", 'Are you sure you want to clear all ticks in the "Selected for Quote" column?') == true) {
    priceList.clearSelectionTicks();
  }
}

function onPriceListExportTicked() {
  trace("onPriceListExport");
  let priceList = new PriceList("PriceList");
  let priceListExport = new PriceListExport("Export");
  priceList.apply(priceListExport);
}

function onPriceListExportSelection() {
  trace("onPriceListExport");
  let priceList = new PriceList("PriceList");
  let priceListExport = new PriceListExport("Export");
  priceList.apply(priceListExport);
}

function onImportPriceList() {
  trace("onImportPriceList");
  let source = PriceList.sheet;
  let destination = SpreadsheetApp.getActiveSpreadsheet();
  source.copyTo(destination).activate();
}


//================================================================================================

class PriceList {
  
  constructor(rangeName = PriceListRangeName) {
    this.range = Range.getByName(rangeName).loadColumnNames();
    trace("NEW " + this.trace);
  }
  
  static get spreadsheet() {
    if (PriceList._spreadsheet === null) {
      PriceList._spreadsheet = SpreadsheetApp.openById(PriceListSpreadsheetId);
    }
    return PriceList._spreadsheet;
  }
  
  static get sheet() {
    if (PriceList._sheet === null) {
      PriceList._sheet = Range.getByName(PriceListRangeName, "", PriceList.spreadsheet).sheet;
    }
    return PriceList._sheet;
  }
  
  update() {
    this.gatherCategories();
  }

  // Method apply
  // Iterate over all price list rows
  //
  apply(handler) {
    trace("PriceList.apply " + this.trace);
    handler.onBegin();
    this.range.forEachRow((range) => {
      const row = new PriceListRow(range);
      if (row.isTitle) {
        handler.onTitle(row);
      } else {
        handler.onRow(row);
      }
    });
    handler.onEnd();
  }

  clearSelectionTicks() {
    trace("PriceList.clearSelectionTicks " + this.trace);
    this.range.forEachRow((range) => {
      let row = new PriceListRow(range);
      if (row.isSelected) {
        row.isSelected = false;
      }
    });
  }
  
  gatherCategories() {
    trace("PriceList.gatherCategories");
    this.range.forEachRow((range) => {
      let row = new PriceListRow(range);
    });
  }

  get trace() {
    return "{PriceList " + this.range.trace + "}";
  }

} // PriceList

PriceList._spreadsheet = null;
PriceList._sheet = null;


//================================================================================================

class PriceListRow extends EventRow {

  constructor(range) {
    super(range);
  }

  get isSelected()      { return this.get("Selected"); } // Accept blank
  set isSelected(value) { this.set("Selected", value); }
  
} // PriceListRow


//================================================================================================

class PriceListExport {
  
  constructor(rangeName) {
    this.range = Range.getByName(rangeName).loadColumnNames();
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
      let exportRow = new EventRow(this.range);
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


