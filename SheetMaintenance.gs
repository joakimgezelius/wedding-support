// Weddings & Events > Templates & Snippets >  > 2026 Wedding Template (USE THIS ONE ONLY)
const templateSheetUrl = "https://docs.google.com/spreadsheets/d/1uBFQrefEIyegbogwe8u0r_QYCcD9CxlWh5BMzwBH3bs/edit?gid=1229250297#gid=1229250297";
const supplierCostingSheetName = "Supplier Costing";
const paramsSheetName = "Params";


function onCleanUpNamedRanges() {
  trace("> onCleanUpNamedRanges");
  SheetMaintenance.cleanUpNamedRanges();
  trace("< onCleanUpNamedRanges");
}

function onInstallSupplierCostingSheet() {
  SheetMaintenance.installSheetTemplate(templateSheetUrl, supplierCostingSheetName);
}

function onInstallParamsSheet() {
  SheetMaintenance.installSheetTemplate(templateSheetUrl, paramsSheetName);
}

class SheetMaintenance {

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

  // Replace a sheet with the corresponding template sheet, and clean up the named ranges
  //
  static installSheetTemplate(templateUrl, sheetName) {
    trace(` > SheetMaintenance.installSheetTemplate ${sheetName}`);
    if (Dialog.confirm("Install New Template Sheet", `Are you sure you want to install a new version of the sheet "${sheetName}"?`) == false) {
      return;
    } 
    const activeSpreadSheet = Spreadsheet.active;
    const templateSpreadSheet = Spreadsheet.openByUrl(templateUrl);
    const templateSheet = templateSpreadSheet.getSheetByName(sheetName);
    const oldSheet = activeSpreadSheet.getSheetByName(sheetName);
    let newSheetTabOrder = 1; // To do: look up detault tab order?
    const newSheet = templateSheet.copyTo(activeSpreadSheet);
    activeSpreadSheet.setActiveSheet(newSheet);
    // NOTE: we're not deleting the old sheet until the new one is in place, as we need to "take over" the global named ranges if they already exist
    newSheet.makeNamedRangesGlobal(); // Take over global named ranges if they already exist - NOTE: must be done before copying the formulas below.
    if (oldSheet !== null) {
      newSheetTabOrder = activeSpreadSheet.getSheetPosition(oldSheet); // Take over the old sheet tab order
      activeSpreadSheet.deleteSheet(oldSheet, true); // false --> prompt for confirmation before delete
    }
    newSheet.name = sheetName;
    // NOTE: After copying, references to external named ranges are broken. To fix, copy the formulas,
    //       however, this must be done after making all named ranges global, in case this is a replacement
    //       copy, see SheetMaintenance.installSheetTemplate() and copyFormulasTo(...) below. If not, the sheet local
    //       named range will take presedence.
    templateSheet.copyFormulasTo(newSheet); // Take a fresh copy of the formulas on the new sheet, as external references would be broken in the sheet copy.
    activeSpreadSheet.nativeSpreadsheet.moveActiveSheet(newSheetTabOrder + 1);
    trace(`< SheetMaintenance.installSheetTemplate ${sheetName}`);
  }

} // class SheetMaintenance
