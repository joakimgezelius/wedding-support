TemplateSpreadsheetId = "1uBFQrefEIyegbogwe8u0r_QYCcD9CxlWh5BMzwBH3bs"; // This is hard-coded to the "2021 Wedding Template" for now

function onApplyFormat() {

  // Identify the template (hard coded reference is OK for now)
  let templateSpreadsheet = Spreadsheet.openById(TemplateSpreadsheetId);

  // Get handle to current sheet (current tab in spreadsheet)
  let activeSheet = Spreadsheet.active.activeSheet;
  let activeRange = activeSheet.fullRange;
  
  // Get handle to same sheet in the template, get it by looking for the same tab name  
  let templateSheet = templateSpreadsheet.getSheetByName(activeSheet.name);
  let templateRange = templateSheet.fullRange;

  Dialog.notify("Applying Template Format...","Please wait this may take a few minutes...");

  // Looping over only Columns to avoid exceeding the execution time which is 360 Seconds.
  // Looping over row by row takes too much time for execution
  for (let column=1; column <= activeSheet.maxColumns; ++column) {
    trace(`Formatting column ${column}`);
    // Getting & Setting Row Width
    // https://developers.google.com/apps-script/reference/spreadsheet/sheet#getColumnWidth(Integer)
    let width = templateSheet.nativeSheet.getColumnWidth(column);
    // https://developers.google.com/apps-script/reference/spreadsheet/sheet#setColumnWidth(Integer,Integer)
    activeSheet.nativeSheet.setColumnWidth(column, width); 

    // Getting & Setting Row Height
    // https://developers.google.com/apps-script/reference/spreadsheet/sheet#getRowHeight(Integer)
    let height = templateSheet.nativeSheet.getRowHeight(column);
    // https://developers.google.com/apps-script/reference/spreadsheet/sheet#setRowHeight(Integer,Integer)  
    activeSheet.nativeSheet.setRowHeight(column,height);

    // Getting & Setting Row Fonts
    // https://developers.google.com/apps-script/reference/spreadsheet/range#getFontFamilies()
    let fonts = templateRange.nativeRange.getFontFamilies();
    // https://developers.google.com/apps-script/reference/spreadsheet/range#setFontFamilies(Object)
    activeRange.nativeRange.setFontFamilies(fonts);

    // Getting & Setting Row Font Size
    // https://developers.google.com/apps-script/reference/spreadsheet/range#getFontSizes()
    let fontSize = templateRange.nativeRange.getFontSizes();
    // https://developers.google.com/apps-script/reference/spreadsheet/range#setFontSizes(Object)
    activeRange.nativeRange.setFontSizes(fontSize); 

    // Getting & Setting Row Font Styles (Italic/Normal)
    // https://developers.google.com/apps-script/reference/spreadsheet/range#getFontStyles()
    let fontStyle = templateRange.nativeRange.getFontStyles();
    // https://developers.google.com/apps-script/reference/spreadsheet/range#setFontStyles(Object)
    activeRange.nativeRange.setFontStyles(fontStyle);

    // Getting & Setting Row Font Weight (Normal/Bold)
    // https://developers.google.com/apps-script/reference/spreadsheet/range#getFontWeights()
    let fontWeight = templateRange.nativeRange.getFontWeights();
    // https://developers.google.com/apps-script/reference/spreadsheet/range#setFontWeights(Object)
    activeRange.nativeRange.setFontWeights(fontWeight);
    
    // Getting & Setting Row Font lines ('underline', 'line-through', or 'none')
    // https://developers.google.com/apps-script/reference/spreadsheet/range#getFontLines()
    let fontLine = templateRange.nativeRange.getFontLines();
    // https://developers.google.com/apps-script/reference/spreadsheet/range#setFontLines(Object)
    activeRange.nativeRange.setFontLines(fontLine);

    // Getting & Setting Row Font Color
    // https://developers.google.com/apps-script/reference/spreadsheet/range#getFontColors()
    let fontColor = templateRange.nativeRange.getFontColors();
    // https://developers.google.com/apps-script/reference/spreadsheet/range#setFontColors(Object)
    activeRange.nativeRange.setFontColors(fontColor);  

    // Getting & Setting Row Number Formats
    // https://developers.google.com/apps-script/reference/spreadsheet/range#getNumberFormats()
    let numberFormat = templateRange.nativeRange.getNumberFormats();
    // https://developers.google.com/apps-script/reference/spreadsheet/range#setNumberFormats(Object) 
    activeRange.nativeRange.setNumberFormats(numberFormat);

    // Getting & Setting Row Background Colors
    // https://developers.google.com/apps-script/reference/spreadsheet/range#getBackgroundObjects()
    let bgObjects = templateRange.nativeRange.getBackgroundObjects();
    // https://developers.google.com/apps-script/reference/spreadsheet/range#setBackgroundObjects(Object)
    activeRange.nativeRange.setBackgroundObjects(bgObjects);  

    // Getting & Setting Row Alignments
    // https://developers.google.com/apps-script/reference/spreadsheet/range#getHorizontalAlignments()
    let hAlignment = templateRange.nativeRange.getHorizontalAlignments();
    // https://developers.google.com/apps-script/reference/spreadsheet/range#setHorizontalAlignments(Object)
    activeRange.nativeRange.setHorizontalAlignments(hAlignment);
  }   
}

function onFormatCoordinator() {
  trace("onFormatCoordinator");
  let eventDetails = new EventDetails();
  let eventDetailsFormater = new EventDetailsFormater(false);
  eventDetails.apply(eventDetailsFormater);

  // https://developers.google.com/apps-script/reference/spreadsheet/range#shiftRowGroupDepth(Integer)  
  // range.nativeRange.shiftRowGroupDepth(-3);          // Removes all the grouping depth in range EventDetails by Delta -3
  // range.nativeRange.shiftRowGroupDepth(1);           // Creates grouping depth for found items range by Delta 1
  
  /*range.forEachRow((range) => {
      const row = new EventRow(range);
      if (row.isTitle) {
       
      } 
    });*/
  

  let sourceSheet = SpreadsheetApp.openById(TemplateSpreadsheetId).getSheetByName("Coordinator");
  let targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Coordinator");
  targetSheet.clearConditionalFormatRules()                       // Removes all the conditional formatting rules from the sheet
  //target.clearFormats();                                        // Clears the sheet of formatting, while preserving contents.
  let rules = sourceSheet.getConditionalFormatRules();            // Gets all the conditional formatting rules from the sheet
  targetSheet.setConditionalFormatRules(rules);                   // Adds new conditional formatting rules
  
  
  /*let sheetValues = source.getDataRange().getValues();
  let sheetBG     = source.getDataRange().getBackgrounds();
  let sheetFC     = source.getDataRange().getFontColors();
  let sheetFF     = source.getDataRange().getFontFamilies();
  let sheetFL     = source.getDataRange().getFontLines();
  let sheetFFa    = source.getDataRange().getFontFamilies();
  let sheetFSz    = source.getDataRange().getFontSizes();
  let sheetFSt    = source.getDataRange().getFontStyles();
  let sheetFW     = source.getDataRange().getFontWeights();
  let sheetHA     = source.getDataRange().getHorizontalAlignments();
  let sheetVA     = source.getDataRange().getVerticalAlignments();
  let sheetNF     = source.getDataRange().getNumberFormats();
  let sheetWR     = source.getDataRange().getWraps();
  let sheetTR     = source.getDataRange().getTextRotations();
  let sheetDir    = source.getDataRange().getTextDirections();
  let sheetNotes  = source.getDataRange().getNotes();

  target.getRange(1,1,sheetValues.length,sheetValues[0].length)
    .setBackgrounds(sheetBG)
    .setFontColors(sheetFC)
    .setFontFamilies(sheetFF)
    .setFontLines(sheetFL)
    .setFontFamilies(sheetFFa)
    .setFontSizes(sheetFSz)
    .setFontStyles(sheetFSt)
    .setFontWeights(sheetFW)
    .setHorizontalAlignments(sheetHA)
    .setVerticalAlignments(sheetVA)
    .setNumberFormats(sheetNF)
    .setWraps(sheetWR)
    .setTextRotations(sheetTR)
    .setTextDirections(sheetDir)
    .setNotes(sheetNotes);*/

  //let width = source.getColumnWidth(sheetValues.length);
  //target.setColumnWidth(sheetValues.length,width);*/




  /*let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Coordinator");
  let numRows = sheet.getMaxRows();
  let range = sheet.getRange(5,1,numRows,36);
  //let address = range.getA1Notation();
  sheet.clearConditionalFormatRules();                                // Clear all rules of CF

  let ruleTitle = SpreadsheetApp.newConditionalFormatRule()           // If the row is title row
      .whenFormulaSatisfied('=$F5="TITLE"')
      .setFontColor('#FFFFFF')
      .setBackground('#666666')
      .setRanges([range])
      .build();
  let ruleTitleRow = sheet.getConditionalFormatRules();
  ruleTitleRow.push(ruleTitle);   
  sheet.setConditionalFormatRules(ruleTitleRow);

  let ruleRow = SpreadsheetApp.newConditionalFormatRule()           // If the row is not title row
      .whenFormulaSatisfied('=$F5<>"TITLE"')
      .setFontColor('#666666')
      .setBackground('#FFFFFF')
      .setRanges([range])
      .build();
  let ruleItemRow = sheet.getConditionalFormatRules();
  ruleItemRow.push(ruleRow);
  sheet.setConditionalFormatRules(ruleItemRow);

  let invoiceAdd = SpreadsheetApp.newConditionalFormatRule()           
      .whenTextContains("ADD")
      .setFontColor('#FFFFFF')
      .setBackground('#FF0000')
      .setRanges([range])
      .build();
  let ruleInvoiceAdd = sheet.getConditionalFormatRules();
  ruleInvoiceAdd.push(invoiceAdd);
  sheet.setConditionalFormatRules(ruleInvoiceAdd);

  let invoiceDone = SpreadsheetApp.newConditionalFormatRule()           
      .whenTextContains("DONE")
      .setFontColor('#666666')
      .setBackground('#B7E1CD')
      .setRanges([range])
      .build();
  let ruleInvoiceDone = sheet.getConditionalFormatRules();
  ruleInvoiceDone.push(invoiceDone);
  sheet.setConditionalFormatRules(ruleInvoiceDone);

  let invoiceNR = SpreadsheetApp.newConditionalFormatRule()          
      .whenTextContains("N/R")
      .setFontColor('#666666')
      .setBackground('#FFFFFF')
      .setRanges([range])
      .build();
  let ruleInvoiceNR = sheet.getConditionalFormatRules();
  ruleInvoiceNR.push(invoiceNR);
  sheet.setConditionalFormatRules(ruleInvoiceNR);

  let currGBP = SpreadsheetApp.newConditionalFormatRule()          
      .whenTextContains("GBP")
      .setFontColor('#666666')
      .setBackground('#B7E1CD')
      .setRanges([range])
      .build();
  let ruleGBP = sheet.getConditionalFormatRules();
  ruleGBP.push(currGBP);
  sheet.setConditionalFormatRules(ruleGBP);

  let currEUR = SpreadsheetApp.newConditionalFormatRule()          
      .whenTextContains("EUR")
      .setFontColor('#FFFFFF')
      .setBackground('#999999')
      .setRanges([range])
      .build();
  let ruleEUR = sheet.getConditionalFormatRules();
  ruleEUR.push(currEUR);
  sheet.setConditionalFormatRules(ruleEUR);*/



  //experiment to get all rules and store in variable and then remove those all and concat new rules with existingrules and clear all rules and add again
  /*let conditionalFormatRules = sheet.getConditionalFormatRules();
  let existingRules = sheet.getConditionalFormatRules();
  let removedRules = [];
  let index = 0;
  let ruleLength = existingRules.length;
  for (index; index < ruleLength; ++index) {
    let ranges = conditionalFormatRules[index].getRanges();
    let j = 0;
    let rangeLength = ranges.length;
    for (j; j < rangeLength; ++j) {
      if (ranges[j].getA1Notation() == address) {
        removedRules.push(existingRules[index]);
      }
    }
  }
  let newRules = [];
  let allRules = existingRules.concat(newRules);
  //clear all rules first and then add again
  sheet.clearConditionalFormatRules(); 
  sheet.setConditionalFormatRules(allRules);*/

}


//=============================================================================================
// Class EventDetailsFormater
//

class EventDetailsFormater {
  constructor(forced) {
    this.forced = forced;
    this.range = Range.getByName("EventDetails","Coordinator");    
    trace("NEW " + this.trace);
  }

  onBegin() {
    this.itemNo = 0;
    this.sectionNo = 0;
    this.eurGbpRate = Range.getByName("EURGBP").value;
    trace("EventDetailsFormater.onBegin - EURGBP=" + this.eurGbpRate);
  }
  
  onEnd() {
    trace("EventDetailsFormater.onEnd - no-op");
  }

  finalizeSectionFormatting() {
    if (this.sectionNo > 1) {// This is not the first title row
     
    }
  }
  
  onTitle(row) {
    trace("EventDetailsFormater.onTitle " + row.itemNo + " " + row.title);  
  }

  onRow(row) {
    ++this.itemNo;
    trace("EventDetailsFormater.onRow " + row.itemNo);
  }
  
  get trace() {
    return `{EventDetailsFormater forced=${this.forced}}`;
  }
}


