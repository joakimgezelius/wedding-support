function onUpdateCoordinator() {
  trace("onUpdateCoordinator");
  if (Dialog.confirm("Update Coordinator - Confirmation Required", "Are you sure you want to update the coordinator? It will overwrite the row numbers, make sure the sheet is sorted properly!") == true) {
    let eventDetailsIterator = new EventDetailsIterator();
    let eventDetailsUpdater = new EventDetailsUpdater();
    eventDetailsIterator.iterate(eventDetailsUpdater);
  }
}

function onCheckCoordinator() {
  trace("onCheckCoordinator");
  let eventDetailsIterator = new EventDetailsIterator();
  let eventDetailsChecker = new EventDetailsChecker();
  eventDetailsIterator.iterate(eventDetailsChecker);
}


//=============================================================================================
// Class EventDetailsIterator
//
class EventDetailsIterator {
  constructor() {
    this.sourceRange = CRange.getByName("EventDetails");
    let columnNamesRange = CRange.getByName("EventDetailsColumnIds");
    this.rowCount = this.sourceRange.height;
    this.data = this.sourceRange.values;
    // let columnNamesRange = this.range.offset(-1, 0, 1); // We expect to find the column names in the row above the data
    EventRow.columnNames = columnNamesRange.values[0];
    trace("NEW " + this.trace);
    trace("Columns: " + EventRow.columnNames);
    for (var column = 0; column < EventRow.columnNames.length; ++column) {
      let columnName = EventRow.columnNames[column];
      if (columnName !== "") {
        EventRow[columnName] = column;
//      trace("Column: " + columnName);
      }
    }
  }

  // Method iterate
  // Iterate over all event rows
  //
  iterate(handler) {
    trace("EventDetailsIterator.iterate " + this.trace);
    handler.onBegin();
    for (var rowOffset = 0; rowOffset < this.rowCount; rowOffset++) {
      var rowRange = this.sourceRange.range.offset(rowOffset, 0, 1);
      var row = new EventRow(this.data[rowOffset], rowOffset, rowRange);
      if (row.isTitle) {
        handler.onTitle(row);
      } else {
        handler.onRow(row);
      }
    }  
    handler.onEnd();
  }
  
  sortByTime() {
    compareTime = function(row1, row2) {
      eventRow1 = new EventRow(row1);
      eventRow2 = new EventRow(row2);
      return eventRow1.compareTime(eventRow2);
    }
    trace("EventDetailsIterator.sortByTime " + this.trace);
    this.data.sort(compareTime);
  }
  
  get trace() {
    return `{EventDetailsIterator range=${this.sourceRange.trace} rowCount=${this.rowCount}`;
  }
}

//=============================================================================================
// Class EventRow
//
class EventRow {
  constructor(data, offset, range) {
    this.data = data;
    this.offset = offset;
    this.range = range;
  }

  getColumnNo(fieldName) { // Private helper
    if (!EventRow.hasOwnProperty(fieldName)) {
      Error.fatal(`Unknown EventRow column: ${fieldName}`);
    }
    let columnNo = EventRow[fieldName];
    return columnNo;
  }
  
  get(fieldName) {
    return this.data[this.getColumnNo(fieldName)];
  }

  getCell(fieldName) { 
    let columnNo = this.getColumnNo(fieldName);
    let cell = this.range.offset(0, columnNo, 1, 1);
    trace(`EventRow.getCell --> ${CRange.trace(cell)}`);
    return cell;
  }

  set(fieldName, value) { 
    let cell = this.getCell(fieldName);
    //trace("EventRow.set " + CRange.trace(cell) + " = " + value);
    cell.setValue(value);
  }

  getA1Notation(fieldName) { 
    let columnNo = this.getColumnNo(fieldName);
    let cell = this.range.offset(0, columnNo, 1, 1);
    return cell.getA1Notation();
  }

  get sectionNo()         { return this.get("SectionNo"); }
  get itemNo()            { return this.get("ItemNo"); }
  get isDecorTicked()     { return this.get("DecorTicked"); }
  get isSupplierTicked()  { return this.get("SupplierTicked"); }
  get isItineraryTicked() { return this.get("ItineraryTicked"); }
  get who()               { return this.get("Who"); }
  get category()          { return this.get("Category"); }
  get status()            { return this.get("Status"); }
  get supplier()          { return this.get("Supplier"); }
  get isTitle()           { return this.category === "Title"; }  // Is this a title row?
  get title()             { return this.get("Description"); }
  get date()              { return this.get("Date"); }
  get time()              { return this.get("Time"); }
  get startTime()         { return this.get("Time"); }
  get endTime()           { return this.get("EndTime"); }
  get location()          { return this.get("Location"); }
  get description()       { return this.get("Description"); }
  get currency()          { return this.get("Currency"); }
  get currencySymbol()    { return this.get("Currency") === GBP ? "£" : "€"; }
  get quantity()          { return this.get("Quantity"); }
  get nativeUnitCost()    { return this.get("NativeUnitCost");  }
  get markup()            { return this.get("Markup"); }
  get unitPrice()         { return this.get("UnitPrice"); }
  get totalPrice()        { return this.get("TotalPrice"); }
  get itemNotes()         { return this.get("ItemNotes"); }
  get notes()             { return this.get("ItemNotes"); }
  get clientNotes()       { return ""; } // this.get("ItemNotes"); }
  get inventoryNotes()    { return ""; } // this.get("ItemNotes"); }
  get links()             { return this.get("Links"); }
  
  compareTime(other) {
    let result = 0;
    if (this.getDate() < other.getDate()) result = -1;
    else if (this.getDate() > other.getDate()) result = 1;
    // Same date, now compare times
    else if (this.getTime() < other.getTime()) result = -1;
    else if (this.getTime() > other.getTime()) return 1;
    else result = 0; // Both date & time are the same
    trace("EventRow.compareTime " + result);
    return result;
  }
}


//=============================================================================================
// Class EventDetailsUpdater
//

class EventDetailsUpdater {
  constructor() {
    trace("NEW " + this.trace);
  }

  onBegin() {
    this.itemNo = 0;
    this.sectionNo = 0;
    this.eurGbpRate = CRange.getByName("EURGBP").value;
    trace("EventDetailsUpdater.onBegin - EURGBP=" + this.eurGbpRate);
  }
  
  onEnd() {
    trace("EventDetailsUpdater.onEnd - no-op");
  }

  onTitle(row) {
    trace("EventDetailsUpdater.onTitle " + row.title);
    this.itemNo = 0;
    ++this.sectionNo;
    if (row.sectionNo === "") { // Only set section id if empty
      row.set("SectionNo", this.generateSectionNo());
    }
    row.set("ItemNo", this.generateSectionNo());
  }

  onRow(row) {
    ++this.itemNo;
    trace("EventDetailsUpdater.onRow " + this.itemNo);
    var currencyA1 = row.getA1Notation("Currency");
    var quantityA1 = row.getA1Notation("Quantity");
    var nativeUnitCostA1 = row.getA1Notation("NativeUnitCost");
    var unitCostA1 = row.getA1Notation("UnitCost");
    var markupA1 = row.getA1Notation("Markup");
    var commissionPercentageA1 = row.getA1Notation("CommissionPercentage");
    var unitPriceA1 = row.getA1Notation("UnitPrice");
    //
    // Set formulas:
    if (row.sectionNo === "") { // Only set section id if empty
      row.set("SectionNo", this.generateSectionNo());
    }
    row.set("ItemNo", this.generateItemNo());
    row.set("UnitCost", Utilities.formatString('=IF(OR(%s="", %s="", %s=0), "", IF(%s="GBP", %s, %s / EURGBP))', currencyA1, nativeUnitCostA1, nativeUnitCostA1, currencyA1, nativeUnitCostA1, nativeUnitCostA1));
    row.set("TotalCost", Utilities.formatString('=IF(OR(%s="", %s=0, %s="", %s=0), "", %s * %s * (1-%s))', quantityA1, quantityA1, unitCostA1, unitCostA1, quantityA1, unitCostA1, commissionPercentageA1));
//  if (row.markup === "") { // Only set markup if empty
//    row.set("Markup", Utilities.formatString('=IF(OR(%s="", %s=0, %s="", %s=0), "", (%s-%s)/%s)', unitCostA1, unitCostA1, unitPriceA1, unitPriceA1, unitPriceA1, unitCostA1, unitCostA1));
//  }
    row.set("UnitPrice", Utilities.formatString('=IF(OR(%s="", %s=0), "", %s * ( 1 + %s))', unitCostA1, unitCostA1, unitCostA1, markupA1));
    row.set("TotalPrice", Utilities.formatString('=IF(OR(%s="", %s=0, %s="", %s=0), "", %s * %s)', quantityA1, quantityA1, unitPriceA1, unitPriceA1, quantityA1, unitPriceA1));
//  row.set("Commission", Utilities.formatString('=IF(OR(%s="", %s=0), "", %s * %s * %s)', commissionPercentageA1, commissionPercentageA1, quantityA1, unitCostA1, commissionPercentageA1));
//  if (=((hour(K215)*60+minute(K215))-(hour(J215)*60+minute(J215)))/60) 
  }

  generateSectionNo() {
    return Utilities.formatString('#%02d', this.sectionNo);
  }

  generateItemNo() {
    if (this.itemNo == 0) 
      return this.generateSectionNo();
    else
      return Utilities.formatString('%s-%02d', this.generateSectionNo(), this.itemNo);
  }

  get trace() {
    return "{EventDetailsUpdater}";
  }
}


//=============================================================================================
// Class EventDetailsChecker
//

class EventDetailsChecker {
  constructor() {
    trace(`NEW ${this.trace}`);
  }
  
  onBegin() {
    trace("EventDetailsChecker.onBegin - no-op");
  }
  
  onEnd() {
    trace("EventDetailsChecker.onEnd - no-op");
  }

  onTitle(row) {
    trace(`EventDetailsChecker.onTitle ${row.title}`);
  }

  onRow(row) {
    trace("EventDetailsChecker.onRow ");
  }
  
  get trace() {
    return "{EventDetailsChecker}";
  }
}
