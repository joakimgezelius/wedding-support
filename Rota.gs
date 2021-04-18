function onUpdateRota() {
  trace("onUpdateRota");
  let aciveSheet = Spreadsheet.active.activeSheet;
  // Pick up range of client sheet references (D7:Dnn)
  let clientSheetListRange = aciveSheet.getRangeByName("ClientSheetList");
  let dataset = "";
  clientSheetListRange.values.forEach(row => { // Add to dataset if not blank
    dataset = dataset + (row[0] != "" ? (dataset != "" ? ";" : "") + `importrange("${row[0]}", "Staff Itinerary!StaffItinerary")` : "") 
  });
  trace(dataset);
  let formula = `=QUERY({${dataset}}, "SELECT * WHERE Col1 IS NOT NULL ORDER BY Col4,Col5,Col6",0)`;
  let cell = aciveSheet.getRangeByName("ImportedActivitiesQuery");
  // https://developers.google.com/apps-script/reference/spreadsheet/range#setValue(Object)
  cell.nativeRange.setValue(formula);
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
