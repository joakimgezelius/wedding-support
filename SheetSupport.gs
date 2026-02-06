var EURTOGBP = 1.11

function GBP(amount, currency) {
  // googlefinance(("CURRENCY:GBPEUR"))
  if (!isFinite(amount)) {
    return "";
  }
  switch (currency) {
    case "EUR":
      gbpAmount = amount / EURTOGBP;
      break;
    case "GBP":
      gbpAmount = amount;
      break;
    case "":
      gbpAmount = "";
      break;
    default:
      gbpAmount = "Unknown Currency";
      break;
    }
  return gbpAmount;
}

/**
 * Calculate marked-up amount. Returns blank if supplied amount or percentage is invalid.
 *
 * @param {number} base amount to be marked up
 * @param {number} mark-up percentage
 * @return {number} the marked-up amount
 */
function MARKUP(amount, percentage) {
  if (!isFinite(amount) || !isFinite(percentage) || amount == 0) {
    return "";
  }
  return amount * (1 + percentage);
}

/**
 * Calculate marked-down amount. Returns blank if supplied amount or percentage is invalid.
 *
 * @param {number} base amount to be marked down
 * @param {number} mark-down percentage
 * @return {number} the marked-down amount
 */
function MARKDOWN(amount, percentage) {
  if (!isFinite(amount) || !isFinite(percentage) || amount == 0) {
    return "";
  }
  return amount * (1 - percentage);
}

function HIDEZERO(value) {
  return (value == 0 ? "" : value);
}

function onSetColour() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var cell = spreadsheet.getActiveSheet().getActiveCell()
  cell.setBackground(cell.getValue());
  trace("onSetColour " + cell.getValue());
}

/**
 * Forces all formulas in the active sheet to recalculate.
 * It does this by finding every "=" in formulas and replacing it with itself.
 */
function onRecalculateFormulas() {
  var sheet = SpreadsheetApp.getActiveSheet(); // https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet-app#getactivesheet
  // This uses the TextFinder API which is much faster than looping through cells
  sheet.createTextFinder("=")
       .matchFormulaText(true)  // Look inside formulas only
       .replaceAllWith("=");    // Replace with itself
       
  SpreadsheetApp.flush(); // Applies all pending spreadsheet changes
  SpreadsheetApp.getUi().alert('Formulas have been refreshed.');
}
