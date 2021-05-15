// For the Menu Update Rota

function onUpdateRota() {
  trace("onUpdateRota");
  let clientSheetList = new ClientSheetList;  
  let dataset = clientSheetList.generateDataset("Staff Itinerary!StaffItinerary", "SELECT * WHERE Col1 IS NOT NULL ORDER BY Col4,Col5,Col6", 0);
  let formula = `=${dataset}`;
  // https://developers.google.com/apps-script/reference/spreadsheet/range#setValue(Object)
  Range.getByName("ImportedActivitiesQuery").nativeRange.setValue(formula);
}

// For the Menu Update Things-to-Buy

function onUpdateThingsToBuy() {
  trace("onUpdateThingsToBuy");
  let clientSheetList = new ClientSheetList;
  let dataset = clientSheetList.generateDataset("EventDetails", "SELECT '${eventName}',Col1,Col6,Col7,Col8,Col9,Col10,Col16 WHERE Col6='Stock to Buy'", 1);
  let formula = `=QUERY(${dataset}, "SELECT * WHERE Col2<>'#01'", 0)`;
  // https://developers.google.com/apps-script/reference/spreadsheet/range#setValue(Object)
  Range.getByName("ImportedThingstoBuyQuery").nativeRange.setValue(formula);
}

// For the Menu Update Things-in-Store

function onUpdateThingsInStore() {
  trace("onUpdateThingsInStore");
  let clientSheetList = new ClientSheetList;
  let dataset = clientSheetList.generateDataset("EventDetails", "SELECT '${eventName}',Col1,Col6,Col7,Col8,Col9,Col10,Col16 WHERE Col6='In Stock'", 1);
  let formula = `=QUERY(${dataset}, "SELECT * WHERE Col2<>'#01'", 0)`;
  // https://developers.google.com/apps-script/reference/spreadsheet/range#setValue(Object)
  Range.getByName("ImportedThingsinStoreQuery").nativeRange.setValue(formula);
}

// For the Menu Update Transportation

function onUpdateTransportation(){
  trace("onUpdateTransportation");
  let clientSheetList = new ClientSheetList;
  let dataset = clientSheetList.generateDataset("EventDetails", "SELECT '${eventName}',Col1,Col6,Col7,Col8,Col9,Col10,Col16 WHERE Col6='Transport'", 1);
  let formula = `=QUERY(${dataset}, "SELECT * WHERE Col2<>'#01'", 0)`;
  // https://developers.google.com/apps-script/reference/spreadsheet/range#setValue(Object)
  Range.getByName("ImportedTransportationQuery").nativeRange.setValue(formula);
}

// For the Menu Update Consumable

function onUpdateConsumable(){
  trace("onUpdateConsumable");
  let clientSheetList = new ClientSheetList;
  let dataset = clientSheetList.generateDataset("EventDetails", "SELECT '${eventName}',Col1,Col6,Col7,Col8,Col9,Col10,Col16 WHERE Col6='Consumable'", 1);
  let formula = `=QUERY(${dataset}, "SELECT * WHERE Col2<>'#01'", 0)`;
  // https://developers.google.com/apps-script/reference/spreadsheet/range#setValue(Object)
  Range.getByName("ImportedConsumableQuery").nativeRange.setValue(formula);
}

// For the Menu Update Shop

function onUpdateShop(){
  trace("onUpdateShop");
  let clientSheetList = new ClientSheetList;
  let dataset = clientSheetList.generateDataset("EventDetails", "SELECT '${eventName}',Col1,Col6,Col7,Col8,Col9,Col10,Col16 WHERE Col6='Shop'", 1);
  let formula = `=QUERY(${dataset}, "SELECT * WHERE Col2<>'#01'", 0)`;
  // https://developers.google.com/apps-script/reference/spreadsheet/range#setValue(Object)
  Range.getByName("ImportedShopQuery").nativeRange.setValue(formula);
}

const ClientSheetListRangeName = "ClientSheetList";

class ClientSheetList {

  constructor(rangeName = ClientSheetListRangeName) {
    this.clientSheetListRange = Range.getByName(rangeName);
    trace("NEW " + this.trace);
  }
  
  generateDataset(rangeName, query, headers) {
    let dataset = "";
    this.clientSheetListRange.values.forEach(row => { // Add to dataset if not blank
      dataset = dataset + (row[0] != "" ? (dataset != "" ? ";" : "") + `QUERY(importrange("${row[0]}", "${rangeName}"), "${query.replace('${eventName}',row[1])}", ${headers})`: "")
    });
    dataset = `{${dataset}}`;
    trace(`ClientSheetList.generateDataset() -> ${dataset}`);
    return dataset;
  }

  get trace() {
    return `{ClientSheetList ${this.clientSheetListRange.trace}}`;
  }

}

//=============================================================================================
// Class Rota
//=============================================================================================

class Rota {

  onActivityColouring() {
    trace("Rota.onActivityColouring ");
  }

  onSupplierColouring() {
    trace("Rota.onSupplierColouring ");
  }

  onLocationColouring () {
    trace("Rota.onLocationColouring ");
  }

  onPerformMagic() {
    trace("Rota.onPerformMagic ");
  }

  static onUpdateRota() {
    trace("Rota.onUpdateRota ");
    //if (Dialog.confirm("Update Rota - Confirmation Required", "Are you sure you want to update the rota? It will overwrite the row numbers, make sure the sheet is sorted properly!") == true) {
    //  trace("Rota.onUpdateRota ");
    //}
  }

  get trace() {
    return "{Rota}";
  }

}
