// For the Menu Update Rota

function onRotaSheetPeriodChanged() {
  trace("onRotaSheetPeriodChanged");
  Dialog.notify("Period Changed", "Sheet will be recalculated, this may take a few seconds...");
  onUpdateRotaSheet();
}

function onUpdateRotaSheet() {
  trace("onUpdateRotaSheet");
  let clientSheetList = new ClientSheetList; 
  clientSheetList.setQuery("RotaQuery",
    "SELECT '${eventName}',Col1,Col6,Col8,Col9,Col10,Col7,Col11,Col12,Col13,Col16 WHERE Col6='Transport' OR  Col6='Rota'", //Col13 IS NOT NULL", 
    "SELECT * WHERE Col2<>'#01' AND NOT LOWER(Col7) CONTAINS 'cancelled' ORDER BY Col4,Col5,Col6");
   
  //let dataset = clientSheetList.generateDataset("Staff Itinerary!StaffItinerary", "SELECT * WHERE Col1 IS NOT NULL ORDER BY Col4,Col5,Col6");
  //let formula = `=${dataset}`;
  // https://developers.google.com/apps-script/reference/spreadsheet/range#setValue(Object)
  //Range.getByName("RotaQuery").nativeRange.setValue(formula);
}

function onCoordinationSheetPeriodChanged() {
  trace("onCoordinationSheetPeriodChanged");
  Dialog.notify("Period Changed", "Sheet will be recalculated, this may take a few seconds...");
  onUpdateCoordinationSheet();
}

function onUpdateCoordinationSheet() {
  trace("onUpdateCoordinationSheet");
  let clientSheetList = new ClientSheetList;
  clientSheetList.setQuery("ThingstoOrderQuery",
    "SELECT '${eventName}',Col1,Col6,Col7,Col8,Col11,Col12,Col13,Col17,Col16 WHERE Col7='To Order' OR Col7='Ordered'", 
    "SELECT * WHERE Col2<>'#01' AND NOT LOWER(Col7) CONTAINS 'cancelled' ORDER BY Col5");
  clientSheetList.setQuery("ThingstoBuyQuery",
    "SELECT '${eventName}',Col1,Col6,Col7,Col8,Col11,Col12,Col13,Col17,Col16 WHERE Col7='To Buy'", 
    "SELECT * WHERE Col2<>'#01' AND NOT LOWER(Col7) CONTAINS 'cancelled' ORDER BY Col5");
  clientSheetList.setQuery("TransportationQuery",
    "SELECT '${eventName}',Col1,Col6,Col8,Col9,Col10,Col7,Col11,Col12,Col13,Col16 WHERE Col6='Transport'", 
    "SELECT * WHERE Col2<>'#01' AND NOT LOWER(Col7) CONTAINS 'cancelled' ORDER BY Col4,Col5");
  clientSheetList.setQuery("ServicesQuery",
    "SELECT '${eventName}',Col1,Col6,Col8,Col9,Col10,Col7,Col11,Col12,Col13,Col16 WHERE Col6='Service'", 
    "SELECT * WHERE Col2<>'#01' AND NOT LOWER(Col7) CONTAINS 'cancelled' AND NOT LOWER(Col7) CONTAINS 'booked' AND NOT LOWER(Col7) CONTAINS 'confirmed' AND NOT LOWER(Col7) CONTAINS 'own arrangement' ORDER BY Col1,Col4,Col5");
  clientSheetList.setQuery("ThingsInStoreQuery",
    "SELECT '${eventName}',Col1,Col6,Col7,Col8,Col11,Col12,Col13,Col17,Col16 WHERE Col7 CONTAINS 'Spain'", 
    "SELECT * WHERE Col2<>'#01' AND NOT LOWER(Col7) CONTAINS 'cancelled' ORDER BY Col5");
  clientSheetList.setQuery("ThingsInShopQuery",
    "SELECT '${eventName}',Col1,Col6,Col7,Col8,Col11,Col12,Col13,Col17,Col16 WHERE Col7 CONTAINS 'Shop'", 
    "SELECT * WHERE Col2<>'#01' AND NOT LOWER(Col7) CONTAINS 'cancelled' ORDER BY Col5");
  clientSheetList.setQuery("RotaQuery",
    "SELECT '${eventName}',Col1,Col6,Col8,Col9,Col10,Col7,Col11,Col12,Col13,Col16 WHERE Col6='Transport' OR Col6='Rota' OR Col6='To Buy'", //Col13 IS NOT NULL", 
    "SELECT * WHERE Col2<>'#01' AND NOT LOWER(Col7) CONTAINS 'cancelled' ORDER BY Col4,Col5,Col6");
}

const ClientSheetListRangeName = "ClientSheetList";

class ClientSheetList {

  constructor(rangeName = ClientSheetListRangeName) {
    this.range = Range.getByName(rangeName);
    trace("NEW " + this.trace);
  }
  
  generateDataset(sourceRangeName, query) {
    let dataset = "";
    this.range.values.forEach(row => { // Add to dataset if not blank
      // NOTE: We're generating queries that return a header row, to avoid problems with empty result sets
      dataset = dataset + (row[0] != "" ? (dataset != "" ? ";" : "") + `QUERY(importrange("${row[0]}", "${sourceRangeName}"), "${query.replace('${eventName}',row[1])}", 1)`: "")
    });
    dataset = `{${dataset}}`;
    trace(`ClientSheetList.generateDataset() -> ${dataset}`);
    return dataset;
  }

  generateQueryFormula(sourceRangeName, innerQuery, outerQuery) {
    let dataset = this.generateDataset(sourceRangeName, innerQuery);
    let queryFormula = `=IFERROR(QUERY(${dataset}, "${outerQuery}", 0))`;
    trace(`ClientSheetList.generateQueryFormula -> ${queryFormula}`);
    return queryFormula;
  }

  setQuery(queryRangeName, innerQuery, outerQuery) {
    let queryRange = Range.getByName(queryRangeName);
    let queryFormula = this.generateQueryFormula( "EventDetails", innerQuery, outerQuery);
    // https://developers.google.com/apps-script/reference/spreadsheet/range#setValue(Object)
    queryRange.nativeRange.setValue(queryFormula);
  }

  get trace() {
    return `{ClientSheetList ${this.range.trace}}`;
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
