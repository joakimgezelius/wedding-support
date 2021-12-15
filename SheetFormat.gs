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

  // https://developers.google.com/apps-script/reference/spreadsheet/range#shiftRowGroupDepth(Integer)  
  // range.nativeRange.shiftRowGroupDepth(-3);          // Removes all the grouping depth in range EventDetails by Delta -3
  // range.nativeRange.shiftRowGroupDepth(1);           // Creates grouping depth for found items range by Delta 1
  
  /*range.forEachRow((range) => {
      const row = new EventRow(range);
      if (row.isTitle) {
       
      } 
    });*/
  
  let source = SpreadsheetApp.openById(TemplateSpreadsheetId).getSheetByName("Coordinator");
  let target = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Coordinator");
  target.clearConditionalFormatRules();                           // Removes all the conditional formatting rules from the sheet
  target.clearFormats();                                          // Clears the sheet of formatting, while preserving contents.
  let rules = source.getConditionalFormatRules();                 // Gets all the conditional formatting rules from the sheet
  target.setConditionalFormatRules(rules);                        // Adds new conditional formatting rules
  
  
  let sheetValues = source.getDataRange().getValues();
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
    .setNotes(sheetNotes);

  //let width = source.getColumnWidth(sheetValues.length);
  //target.setColumnWidth(sheetValues.length,width);


  /* Adds the rule once on run Format coordinator 
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Coordinator");
  //sheet.clearConditionalFormatRules();                          // Removes all the conditional formatting rules from the sheet
  let numRows = sheet.getLastRow();
  let rangeToFormat = sheet.getRange("A21:AO"+numRows);           // Includes rows in entire sheet excluding First Title row i.e A20
  let rule1 = SpreadsheetApp.newConditionalFormatRule()           // If the row is title row
      .whenFormulaSatisfied('=$F21="TITLE"')
      .setFontColor('#FFFFFF')
      .setBackground('#666666')
      .setRanges([rangeToFormat])
      .build();
  let ruleTitleRow = sheet.getConditionalFormatRules();
  ruleTitleRow.push(rule1);   
  sheet.setConditionalFormatRules(ruleTitleRow);

  let rule2 = SpreadsheetApp.newConditionalFormatRule()           // If the row is not title row
      .whenFormulaSatisfied('=$F21<>"TITLE"')
      .setFontColor('#666666')
      .setBackground('#FFFFFF')
      .setRanges([rangeToFormat])
      .build();
  let ruleItemRow = sheet.getConditionalFormatRules();
  ruleItemRow.push(rule2);
  sheet.setConditionalFormatRules(ruleItemRow); */

  eventDetails.apply(eventDetailsFormater);
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


