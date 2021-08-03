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
  let range = Coordinator.eventDetailsRange;

  // https://developers.google.com/apps-script/reference/spreadsheet/range#shiftRowGroupDepth(Integer)  
  // range.nativeRange.shiftRowGroupDepth(-3);          // Removes all the grouping depth in range EventDetails by Delta -3
  // range.nativeRange.shiftRowGroupDepth(1);           // Creates grouping depth for found items range by Delta 1
  range.forEachRow((range) => {
      const row = new EventRow(range);
      if (row.isTitle) {
       range.nativeRange.shiftRowGroupDepth(1);
      } 
    });

  /*
  let itemRows = 
  for (let i=1; i <= itemRows; i++) {

  }*/

  /*
  const row = new EventRow(range);
  if(row.isTitle) {
    range.setFontColors('#FFFFFF').setFontSizes(10).setBackgroundColor('#666666');
  }*/

  eventDetails.apply(eventDetailsFormater);
}


class EventDetailsFormater {
  constructor(forced) {
    this.forced = forced;
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


