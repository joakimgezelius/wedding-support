
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
        namedRange.remove(false); // false --> prompt for confirmation before delete
      }
      else if (namedRange.name.includes('!')) { // This is a local named range
        Dialog.notify("Local Named Range", `Local named range found: ${namedRange.name}, making it global`);
        namedRange.makeGlobal();
      }
    });
  }

}

class SupplierCostingTemplate {
  // Weddings & Events > Templates & Snippets >  > 2026 Wedding Template (USE THIS ONE ONLY)
  static get templateSheetUrl() { return "https://docs.google.com/spreadsheets/d/1uBFQrefEIyegbogwe8u0r_QYCcD9CxlWh5BMzwBH3bs/edit?gid=1229250297#gid=1229250297"; }
  static get templateSheetName() { return "Supplier Costing"; }

  static install() {
    const activeSpreadSheet = Spreadsheet.active;
    const templateSpreadSheet = Spreadsheet.openByUrl(SupplierCostingTemplate.templateSheetUrl);
    const templateSheet = templateSpreadSheet.getSheetByName(SupplierCostingTemplate.templateSheetName);
    const oldSheet = activeSpreadSheet.getSheetByName(SupplierCostingTemplate.templateSheetName);
    let oldSheetTabOrder = 1; // To do: look up detault tab order?
    if (oldSheet !== null) {
      oldSheetTabOrder = activeSpreadSheet.getSheetPosition(oldSheet);
      activeSpreadSheet.deleteSheet(oldSheet, false); // false --> prompt for confirmation before delete
    }
    const newSheet = templateSheet.copyTo(activeSpreadSheet);
    newSheet.name = SupplierCostingTemplate.templateSheetName;
    activeSpreadSheet.setActiveSheet(newSheet);
    activeSpreadSheet.nativeSpreadsheet.moveActiveSheet(oldSheetTabOrder + 1);
    newSheet.makeNamedRangesGlobal();
  }
}
