
function onInstallSupplierCostingSheet() {
  trace("> onInstallSupplierCostingSheet");
  SupplierCostingTemplate.install();
  trace("< onInstallSupplierCostingSheet");
}

class SheetMaintenance {

  installSupplierCostingSheet() {
    ;
  }

}

class SupplierCostingTemplate {

  static get templateSheetUrl() { return "https://docs.google.com/spreadsheets/d/1k3eS6-9KZ88IubW01tlCzXpSRFW9bfk_xd7lYEnQpso/edit?gid=71237569#gid=71237569"; }
  static get templateSheetName() { return "Supplier Costing"; }

  static install() {
    const templateSpreadSheet = Spreadsheet.openByUrl(SupplierCostingTemplate.templateSheetUrl);
    const templateSheet = templateSpreadSheet.getSheetByName(SupplierCostingTemplate.templateSheetName);
    let activeSpreadSheet = Spreadsheet.active;
    let newSheet = templateSheet.copyTo(activeSpreadSheet);
    let oldSheet = activeSpreadSheet.getSheetByName(SupplierCostingTemplate.templateSheetName);
    const oldSheetTabOrder = activeSpreadSheet.getSheetPosition(oldSheet);
    activeSpreadSheet.setActiveSheet(newSheet);
    activeSpreadSheet.deleteSheet(oldSheet);
    activeSpreadSheet.nativeSpreadsheet.moveActiveSheet(oldSheetTabOrder + 1);
    //.activate();
    // SupplierTotalCost
    // SupplierCosting
//    newSheet.moveto();
  }
}
