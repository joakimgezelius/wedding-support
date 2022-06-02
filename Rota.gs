//----------------------------------------------------------------------------------------
// Rota Sheet

function onRotaSheetPeriodChanged() {
  trace("onRotaSheetPeriodChanged");
  Dialog.notify("Period Changed", "Sheet will be recalculated, this may take a few seconds...");
  onUpdateRotaSheet();
}

function onUpdateRotaSheet() {
  trace("onUpdateRotaSheet");
  let clientSheetList = new ClientSheetList;
  clientSheetList.setQuery("RotaQuery",
    "SELECT '${eventName}',Col1,Col6,Col8,Col9,Col10,Col7,Col11,Col12,Col13,Col16 WHERE (Col4=TRUE OR LOWER(Col6) CONTAINS 'transport' OR LOWER(Col6) CONTAINS 'rota')", //Col13 IS NOT NULL", 
    "SELECT * WHERE Col2<>'#01' AND NOT LOWER(Col7) CONTAINS 'cancelled' ORDER BY Col4,Col5,Col6");
}


//----------------------------------------------------------------------------------------
// Event Coordination Sheet - currently only one sheet:
// https://docs.google.com/spreadsheets/d/1BWwA-yRJQyUooIh5nodzTPpqcYxeFHf8-0Uo2V7JQbk/edit?pli=1#gid=233397022

function onEventCoordinationSheetPeriodChanged() {
  trace("onEventCoordinationSheetPeriodChanged");
  Dialog.notify("Period Changed", "Sheet will be recalculated, this may take a few seconds...");
  onUpdateEventCoordinationSheet();
}

function onUpdateEventCoordinationSheet() {
  trace("onUpdateEventCoordinationSheet");
  let clientSheetList = new ClientSheetList;

  clientSheetList.setQuery("ThingsToOrderQuery",
    "SELECT '${eventName}',Col1,Col6,Col7,Col8,Col11,Col12,Col13,Col18,Col16 WHERE LOWER (Col7) CONTAINS 'to order' OR Col7='Ordered'", 
    "SELECT * WHERE Col2<>'#01' AND NOT LOWER(Col7) CONTAINS 'cancelled' ORDER BY Col5");

  clientSheetList.setQuery("ThingsOrderedQuery",
    "SELECT '${eventName}',Col1,Col6,Col7,Col8,Col11,Col12,Col13,Col18,Col16 WHERE Col7='Ordered'", 
    "SELECT * WHERE Col2<>'#01' AND NOT LOWER(Col7) CONTAINS 'cancelled' ORDER BY Col5");

  clientSheetList.setQuery("ThingsToBuyQuery",
    "SELECT '${eventName}',Col1,Col6,Col7,Col8,Col11,Col12,Col13,Col18,Col16 WHERE Col7='To Buy' OR Col7='To Collect'", 
    "SELECT * WHERE Col2<>'#01' AND NOT LOWER(Col7) CONTAINS 'cancelled' ORDER BY Col5");

  clientSheetList.setQuery("TransportationQuery",
    "SELECT '${eventName}',Col1,Col6,Col8,Col9,Col10,Col7,Col11,Col12,Col13,Col16 WHERE Col6='Transport'", 
    "SELECT * WHERE Col2<>'#01' AND NOT LOWER(Col7) CONTAINS 'cancelled' ORDER BY Col5");

  clientSheetList.setQuery("ServicesQuery",
    "SELECT '${eventName}',Col1,Col6,Col8,Col9,Col10,Col7,Col11,Col12,Col13,Col16 WHERE Col6='Service'", 
    "SELECT * WHERE Col2<>'#01' AND NOT LOWER(Col7) CONTAINS 'cancelled' AND NOT LOWER(Col7) CONTAINS 'booked' AND NOT LOWER(Col7) CONTAINS 'confirmed' AND NOT LOWER(Col7) CONTAINS 'own arrangement' ORDER BY Col4,Col5");

  clientSheetList.setQuery("ThingsInStoreQuery",
    "SELECT '${eventName}',Col1,Col6,Col7,Col8,Col11,Col12,Col13,Col18,Col16 WHERE LOWER(Col7) CONTAINS 'spain'", 
    "SELECT * WHERE Col2<>'#01' AND NOT LOWER(Col7) CONTAINS 'cancelled' ORDER BY Col5");

  clientSheetList.setQuery("ThingsInShopQuery",
    "SELECT '${eventName}',Col1,Col6,Col7,Col8,Col11,Col12,Col13,Col18,Col16 WHERE LOWER(Col7) CONTAINS 'shop' or LOWER(Col7) CONTAINS 'gibraltar'",
    "SELECT * WHERE Col2<>'#01' AND NOT LOWER(Col7) CONTAINS 'cancelled' ORDER BY Col5");

  clientSheetList.setQuery("RotaQuery",
    "SELECT '${eventName}',Col1,Col6,Col8,Col9,Col10,Col7,Col11,Col12,Col13,Col16 WHERE (LOWER(Col6) CONTAINS 'transport' OR  LOWER(Col6) CONTAINS 'rota')", // Col4=TRUE OR Col13 IS NOT NULL", 
    "SELECT * WHERE Col2<>'#01' AND NOT LOWER(Col7) CONTAINS 'cancelled' ORDER BY Col4,Col5,Col6");

  clientSheetList.setQuery("HotelReservationsQuery",
    "SELECT '${eventName}',Col1,Col6,Col7,Col8,Col12,Col14,Col18,Col28,Col16 WHERE LOWER(Col6) CONTAINS 'hotel'", 
    "SELECT * WHERE Col2<>'#01' AND NOT LOWER(Col7) CONTAINS 'cancelled' ORDER BY Col5");
}


//----------------------------------------------------------------------------------------
// ClientSheetList

// ClientSheetListRangeName is the name of the named range holding the list of client sheets.
// The first column of this range (i.e. row[0]) holds the URL of the client sheet, if not empty.
// 
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
