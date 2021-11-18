function onUpdatePriceListForced() {
  trace("onUpdatePriceListForced");  
  if (Dialog.confirm("Forced Price List Update - Confirmation Required", "Are you sure you want to force-update the price list? It will overwrite row numbers and formulas, make sure the sheet is sorted properly!") == true) {
    let priceList = new PriceList();
    let priceListUpdater = new PriceListUpdater(true);
    priceList.apply(priceListUpdater);
  }
}

function onRefreshPriceList() {
  trace("onRefreshPriceList");
}

function onUpdatePackages() {
  trace("onUpdatePackages");
//  if (Dialog.confirm("Forced Coordinator Update - Confirmation Required", "Are you sure you want to force-update the coordinator? It will overwrite row numbers and formulas, make sure the sheet is sorted properly!") == true) {
//  }
}

function onPriceListClearExport() {
  trace("onPriceListClearExport");
  let priceListExport = new PriceListExport("Export");
  priceListExport.clear();
}

function onPriceListClearSelectionTicks() {
  trace("onPriceListClearSelectionTicks");
  let priceList = new PriceList();
  if (Dialog.confirm("Clear Selection Ticks", 'Are you sure you want to clear all ticks in the "Selected for Quote" column?') == true) {
    priceList.clearSelectionTicks();
  }
}

function onPriceListExportTicked() {
  trace("onPriceListExport");
  let priceList = new PriceList("PriceList");
  let priceListExport = new PriceListExport("Export");
  priceList.apply(priceListExport);
}

function onPriceListExportSelection() {
  trace("onPriceListExport");
  let priceList = new PriceList("PriceList");
  let priceListExport = new PriceListExport("Export");
  priceList.apply(priceListExport);
}

function onImportPriceList() {
  trace("onImportPriceList");
  let source = PriceList.sheet;
  let destination = SpreadsheetApp.getActiveSpreadsheet();
  source.copyTo(destination).activate();
}


//=============================================================================================
// Class PriceListUpdater
//

class PriceListUpdater {
  
  constructor(forced) {
    this.forced = forced;
    trace("NEW " + this.trace);
  }

  onBegin() {
    this.eurGbpRate = Range.getByName("EURGBP").value;
    trace("PriceListUpdater.onBegin - EURGBP=" + this.eurGbpRate);
  }
  
  onEnd() {
    trace("PriceListUpdater.onEnd - no-op");
  }

  onRow(row) {
    trace("PriceListUpdater.onRow");
    let a1_selected = row.getA1Notation("Selected");
    let a1_quantity = row.getA1Notation("Quantity");
    let a1_currency = row.getA1Notation("Currency");
    let a1_nativeUnitCost = row.getA1Notation("NativeUnitCost");
    let a1_vat = row.getA1Notation("VAT");
    let a1_nativeUnitCostWithVAT = row.getA1Notation("NativeUnitCostWithVAT");
    let a1_unitCost = row.getA1Notation("UnitCost");    
    let a1_markup = row.getA1Notation("Markup");
    let a1_commissionPercentage = row.getA1Notation("CommissionPercentage");
    let a1_unitPrice = row.getA1Notation("UnitPrice");

    this.setNativeUnitCost(row);
    row.nativeUnitCostWithVAT = `=IF(OR(${a1_currency}="", ${a1_nativeUnitCost}="", ${a1_nativeUnitCost}=0), "", ${a1_nativeUnitCost}*(1+${a1_vat}))`;
    row.getCell("NativeUnitCostWithVAT").setNumberFormat(row.currencyFormat);
    row.unitCost = `=IF(OR(${a1_currency}="", ${a1_nativeUnitCostWithVAT}="", ${a1_nativeUnitCostWithVAT}=0), "", IF(${a1_currency}="GBP", ${a1_nativeUnitCostWithVAT}, ${a1_nativeUnitCostWithVAT} / EURGBP))`;
    row.totalGrossCost = `=IF(OR(${a1_quantity}="", ${a1_quantity}=0, ${a1_unitCost}="", ${a1_unitCost}=0, ${a1_selected}=FALSE), "", ${a1_quantity} * ${a1_unitCost})`;
    row.totalNettCost = `=IF(OR(${a1_quantity}="", ${a1_quantity}=0, ${a1_unitCost}="", ${a1_unitCost}=0, ${a1_selected}=FALSE), "", ${a1_quantity} * ${a1_unitCost} * (1-${a1_commissionPercentage}))`;
    row.unitPrice = `=IF(OR(${a1_unitCost}="", ${a1_unitCost}=0), "", ${a1_unitCost} * ( 1 + ${a1_markup}))`;
    row.totalPrice = `=IF(OR(${a1_quantity}="", ${a1_quantity}=0, ${a1_unitPrice}="", ${a1_unitPrice}=0, ${a1_selected}=FALSE), "", ${a1_quantity} * ${a1_unitPrice})`;
//  row.commission = `=IF(OR(${a1_commissionPercentage}="", ${a1_commissionPercentage}=0), "", ${a1_quantity} * ${a1_unitCost} * ${a1_commissionPercentage})`;
    if (row.isInStock && this.forced) { // Set mark-up and commission for in-stock items
      trace("- Set In Stock commision & mark-up on ");
      row.commissionPercentage = 0.5;
      row.markup = 0;
    }
  }

  setNativeUnitCost(row) {
    let budgetUnitCostCell = row.getCell("BudgetUnitCost");
    let a1_budgetUnitCost = budgetUnitCostCell.getA1Notation();
    let nativeUnitCost = String(row.nativeUnitCost);
    let nativeUnitCostCell = row.getCell("NativeUnitCost");
    let currencyFormat = row.currencyFormat;
    //trace(`Cell: ${nativeUnitCostCell.getA1Notation()} ${currencyFormat}`);
    budgetUnitCostCell.setNumberFormat(currencyFormat);
    nativeUnitCostCell.setNumberFormat(currencyFormat);
    if (nativeUnitCost === "") { // Set native unit cost equal to budget unit cost if not set 
      //trace(`Set nativeUnitCost to =${budgetUnitCostA1}`);
      row.nativeUnitCost = `=${a1_budgetUnitCost}`;
      nativeUnitCostCell.setFontColor("#ccccff"); // Make text grey
    } else {
      //trace(`setNativeUnitCost: nativeUnitCost="${nativeUnitCost} - make it black`);
      nativeUnitCostCell.setFontColor("#000000"); // Black
    }
  }

  get trace() {
    return `{PriceListUpdater forced=${this.forced}}`;
  }
}




