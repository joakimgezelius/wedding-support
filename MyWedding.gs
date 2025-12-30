
function onUpdateWeddingTemplate() {
  MyWedding.updateWeddingTemplate();
}


class MyWedding {
/**
 * Main function to update the wedding template using mapping sheet + BigQuery data
 * (Currently uses a sheet named bigquery_data until BigQuery API is added)
 */

  static updateWeddingTemplate() {

    const ss = SpreadsheetApp.getActive();

    // Sheets
    const templateSheet = ss.getSheetByName("Wedding Journal");
    const mappingSheet  = ss.getSheetByName("client_template_mapping_sheet");
    const dataSheet     = ss.getSheetByName("bigquery_data");

    if (!templateSheet || !mappingSheet || !dataSheet) {
      SpreadsheetApp.getUi().alert("Error: One or more required sheets are missing.");
      return;
    }

    // Load mapping (template_row_name, big_query_field_name)
    const mapping = mappingSheet.getRange(
      2, 1,
      mappingSheet.getLastRow() - 1,
      2
    ).getValues();

    // Load BigQuery (or temporary) data
    const rawData = dataSheet.getRange(
      2, 1,
      1,
      dataSheet.getLastColumn()
    ).getValues()[0];

    const header = dataSheet.getRange(
      1, 1,
      1,
      dataSheet.getLastColumn()
    ).getValues()[0];

    // Convert data into an object: { id: "...", dealname: "...", amount: 450 }
    const dataObj = {};
    header.forEach((colName, i) => {
      dataObj[colName] = rawData[i];
    });


    // Loop through mapping and write values into the template
    mapping.forEach(row => {
      const templateLabel = row[0];    // e.g. "Contact Details"
      const bqField       = row[1];    // e.g. "amount"

      const value = dataObj[bqField];

      if (value === undefined || value === "") {
        Logger.log("Skipping: no data for field " + bqField);
        return;
      }

      // Find the label in the wedding template sheet
      const foundCell = templateSheet.createTextFinder(templateLabel).findNext();

      if (!foundCell) {
        Logger.log("Label not found in template: " + templateLabel);
        return;
      }

      // Write in column C (two columns to the right)
      const targetCell = foundCell.offset(0, 2);
      targetCell.setValue(value);

      Logger.log("Updated: " + templateLabel + " → " + value);
    });

    SpreadsheetApp.getUi().alert("Wedding template updated successfully!");
  }

} // MyWedding