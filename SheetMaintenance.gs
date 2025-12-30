
function onInstallSupplierCostingSheet() {
  trace("> onInstallSupplierCostingSheet");
  SupplierCostingTemplate.install();
  trace("< onInstallSupplierCostingSheet");
}

function onCleanUpNamedRanges() {
  trace("> onCleanUpNamedRanges");
  SheetMaintenance.cleanUpNamedRanges();
  trace("< onCleanUpNamedRanges");
}


class SheetMaintenance {

  static installSupplierCostingSheet() {
    ;
  }
  
  static cleanUpNamedRanges() {
    Spreadsheet.active.iterateOverNamedRanges((namedRange) => { // Callback to arrow function
      if (namedRange.range === null) { // This is an invalid range, to be deleted
        trace(`  cleanUpNamedRanges, invalid range, delete: ${namedRange.trace}`);
        //namedRange.remove();
      }
    });
  }

}

class SupplierCostingTemplate {

  static get templateSheetUrl() { return "https://docs.google.com/spreadsheets/d/1k3eS6-9KZ88IubW01tlCzXpSRFW9bfk_xd7lYEnQpso/edit?gid=71237569#gid=71237569"; }
  static get templateSheetName() { return "Supplier Costing"; }

  static install() {
    const activeSpreadSheet = Spreadsheet.active;
    const templateSpreadSheet = Spreadsheet.openByUrl(SupplierCostingTemplate.templateSheetUrl);
    const templateSheet = templateSpreadSheet.getSheetByName(SupplierCostingTemplate.templateSheetName);
    const oldSheet = activeSpreadSheet.getSheetByName(SupplierCostingTemplate.templateSheetName);
    let oldSheetTabOrder = 1; // To do: look up detault tab order?
    if (oldSheet !== null) {
      oldSheetTabOrder = activeSpreadSheet.getSheetPosition(oldSheet);
      activeSpreadSheet.deleteSheet(oldSheet);
    }
    const newSheet = templateSheet.copyTo(activeSpreadSheet);
    newSheet.name = SupplierCostingTemplate.templateSheetName;
    activeSpreadSheet.setActiveSheet(newSheet);
    activeSpreadSheet.nativeSpreadsheet.moveActiveSheet(oldSheetTabOrder + 1);

    // Sort out the named ranges
  }
}
