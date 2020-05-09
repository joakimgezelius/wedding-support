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

function onPriceListExport() {
  PriceList.onExport();
}

class PriceList {
  
  constructor() {
    this.range = Range.getByName("PriceList");
    this.rowCount = this.range.getHeight();
    this.data = this.range.getValues();
    trace("NEW " + this.trace());
  }
  
  static onExport() {
    trace("onExport");
    Dialog.notify("onExport", "onExport");
  }

  update() {
    this.gatherCategories();
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

}


//  eventDetailsIterator.iterate(budgetBuilder);
