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
  //let eventDetails = new EventDetails();
  //let eventDetailsFormater = new EventDetailsFormater(false);
  //eventDetails.apply(eventDetailsFormater);

  // https://developers.google.com/apps-script/reference/spreadsheet/range#shiftRowGroupDepth(Integer)  
  // range.nativeRange.shiftRowGroupDepth(-3);          // Removes all the grouping depth in range EventDetails by Delta -3
  // range.nativeRange.shiftRowGroupDepth(1);           // Creates grouping depth for found items range by Delta 1
  
  /*range.forEachRow((range) => {
      const row = new EventRow(range);
      if (row.isTitle) {
       
      } 
  });*/
  
  // Formmatings with reference from template sheet to Coordinator sheet
  let sourceSheet = SpreadsheetApp.openById(TemplateSpreadsheetId).getSheetByName("Coordinator");
  let targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Coordinator");  
  let range = targetSheet.getRange("A4:AJ"+targetSheet.getLastRow());
  targetSheet.clearConditionalFormatRules();                                      // Removes all the conditional formatting rules from the sheet
  //targetSheet.clearFormats();                                                     // Clears the sheet of formatting, while preserving contents.
  
  // Conditional formatting based on template sheet - Coordinator
  let rules = sourceSheet.getConditionalFormatRules();
  let newRules = [], i = 0;
  
  // Hard-coded formatting rule to specific range & avoid overwriting of Custom Formula rule for TITLE row
  let rangeMarkup = targetSheet.getRange("AA4:AA" + targetSheet.getLastRow());  // For Markup
  let ruleMarkup = SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('%')
      .setBackground("#D9EAD3")
      .setFontColor("#666666")
      .setRanges([rangeMarkup])
      .build();
  newRules.push(ruleMarkup);

  let rangeComm = targetSheet.getRange("AB4:AB" + targetSheet.getLastRow());    // For Commision
  let ruleComm1 = SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('%')
      .setBackground("#CCCCCC")
      .setFontColor("#434343")
      .setRanges([rangeComm])
      .build();
  newRules.push(ruleComm1);
  
  /*let ruleComm2 = SpreadsheetApp.newConditionalFormatRule()         // For commision greater than 1%
      .whenNumberGreaterThan(1)
      .setBackground("#FF0000")
      .setFontColor("#FFFFFF")
      .setRanges([rangeComm])
      .build();
  newRules.push(ruleComm2);*/

  let ruleLength = rules.length;
  for (i; i < ruleLength; ++i) {
    let rule = rules[i];
    let booleanCondition = rule.getBooleanCondition();
    if (booleanCondition != null) {
      let newRule = SpreadsheetApp.newConditionalFormatRule()
        .withCriteria(booleanCondition.getCriteriaType(), booleanCondition.getCriteriaValues())
        .setBackground(booleanCondition.getBackground())
        .setFontColor(booleanCondition.getFontColor())
        .setRanges([range])
        .build();
      newRules.push(newRule);
    }
  }

  targetSheet.setConditionalFormatRules(newRules);                                     
  
  /* Formatting based on template sheet
  let sheetValues = sourceSheet.getDataRange().getValues();
  let sheetBG     = sourceSheet.getDataRange().getBackgrounds();
  let sheetFC     = sourceSheet.getDataRange().getFontColors();
  let sheetFF     = sourceSheet.getDataRange().getFontFamilies();
  let sheetFL     = sourceSheet.getDataRange().getFontLines();
  let sheetFFa    = sourceSheet.getDataRange().getFontFamilies();
  let sheetFSz    = sourceSheet.getDataRange().getFontSizes();
  let sheetFSt    = sourceSheet.getDataRange().getFontStyles();
  let sheetFW     = sourceSheet.getDataRange().getFontWeights();
  let sheetHA     = sourceSheet.getDataRange().getHorizontalAlignments();
  let sheetVA     = sourceSheet.getDataRange().getVerticalAlignments();
  let sheetNF     = sourceSheet.getDataRange().getNumberFormats();
  let sheetWR     = sourceSheet.getDataRange().getWraps();
  let sheetTR     = sourceSheet.getDataRange().getTextRotations();
  let sheetDir    = sourceSheet.getDataRange().getTextDirections();
  let sheetNotes  = sourceSheet.getDataRange().getNotes();

  targetSheet.getRange(1,1,sheetValues.length,sheetValues[0].length)  // "A1:AJ"+targetSheet.getLastRow()
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

  //let width = sourceSheet.getColumnWidth(sheetValues.length);
  //targetSheet.setColumnWidth(sheetValues.length,width);

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
    trace("EventDetailsFormater.onBegin");
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


