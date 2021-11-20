
class PriceList {
  
  /*constructor(rangeName = PriceListRangeName) {
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
  }*/
  
  constructor() {
    trace("constructing PriceList object...");
    this.range = Range.getByName().loadColumnNames();
    this.rowCount = this.range.height;
    trace("NEW " + this.trace);
  }

  // Method apply
  // Iterate over all event rows (using Range Row Iterator), call handler methods 
  //
  apply(handler) {
    trace(`${this.trace}.apply`);
    handler.onBegin();
    this.range.forEachRow((range) => {
      const row = new PriceListRow(range);
      handler.onRow(row);
    });
    handler.onEnd();
  }
  
  get trace() {
    return `{PriceList range=${this.range.trace} rowCount=${this.rowCount}`;
  }
} // PriceList


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


