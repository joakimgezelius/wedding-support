
class Formatting {

  listConditionalFormattingRules() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const rules = sheet.getConditionalFormatRules();
    
    if (rules.length === 0) {
        Logger.log("No conditional formatting rules found on this sheet.");
        return;
    }

    Logger.log("Conditional Formatting Rules for sheet: " + sheet.getName());
    Logger.log("--------------------------------------------------");
    
    rules.forEach((rule, index) => {
        let ranges = rule.getRanges().map(r => r.getA1Notation()).join(', ');
        Logger.log(`Rule ${index + 1}: Applied to Range(s): ${ranges}`);
    });
  }

  removeConditionalFormattingRules() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const rules = sheet.getConditionalFormatRules();
    
    if (rules.length === 0) {
        trace("No conditional formatting rules found on this sheet.");
        return;
    }

    trace("Removing all conditional formatting rules from sheet: " + sheet.getName());
    
    sheet.setConditionalFormatRules([]);
    trace("All conditional formatting rules removed.");
  }

}
