//----------------------------------------------------------------------------------------
// Rota Sheet

function onRotaSheetPeriodChanged() {
  trace("onRotaSheetPeriodChanged");
  Dialog.notify("Period Changed", "Sheet will be recalculated, this may take a few seconds...");
  onEventCoordinationSheetPeriodChanged();
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
  Dialog.notify("Refreshing Details", "Sheet will be recalculated, this may take a few seconds...");
  updateMasterQuery();
}

function updateMasterQuery() {
  trace("updateMasterQuery");
  let clientSheetList = new ClientSheetList;
  // 
  // Note: We now use a two-phased appraoch, first phase collects a master data set based on a query we compile in code (next couple of lines),
  //       after which secodary queries on each tab further filters the data set.
  //
  clientSheetList.setQuery("MasterQuery",
    "SELECT '${eventName}',Col1,Col2,Col3,Col4,Col5,Col6,Col7,Col8,Col9,Col10,Col11,Col12,Col13,Col14,Col15,Col16,Col17,Col18,Col19,Col20,Col21,Col22,Col23,Col24,Col25,Col26,Col27,Col28,Col29,Col30,Col31,Col32,Col33,Col34,Col35", 
    "SELECT * WHERE Col2<>'#01'");
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
      dataset = dataset + (row[0] != "" ? (dataset != "" ? ";" : "") + `IFERROR(QUERY(importrange("${row[0]}", "${sourceRangeName}"), "${query.replace('${eventName}',row[1])}", 1))`: "")
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
