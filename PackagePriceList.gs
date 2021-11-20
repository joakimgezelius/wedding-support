function onUpdatePriceListForced() {
  trace("onUpdatePriceListForced");
  if (Dialog.confirm("Forced Price List Update - Confirmation Required", "Are you sure you want to force-update the price list? It will overwrite row numbers and formulas, make sure the sheet is sorted properly!") == true) {
    const priceListIndex = new PriceListIndex;
    priceListIndex.updatePriceLists();
  }
}

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

//=============================================================================================
// Class PriceListIndex
//

class PriceListIndex {
  
  constructor(indexRangeName = "PriceListIndex") {
    this._range = Range.getByName(indexRangeName).loadColumnNames();
    this._trace = `{PriceListIndex ${this._range.trace}}`;
    trace("NEW " + this.trace);
  }

  // Iterate over all price list categories/sections and apply the forced update to each one
  //
  updatePriceLists() {
    this._range.forEachRow((range) => {
      const row = new RangeRow(range);
      const category = row.get("Category");
      const rangeName = row.get("Range");
      if (category !== "") {
        trace(`updatePriceList: ${category}, ${rangeName}`);
        // Force-update the category/section at hand
        let eventDetails = new EventDetails(rangeName);
        let eventDetailsUpdater = new EventDetailsUpdater(true);
        eventDetails.apply(eventDetailsUpdater);
      }
    })
  }

  get trace() {
    return this._trace;
  }

}
