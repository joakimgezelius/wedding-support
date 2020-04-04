function onUpdateCoordinator() {
  trace("onUpdateCoordinator");
  if (Dialog.confirm("Update Coordinator - Confirmation Required", "Are you sure you want to update the coordinator? It will overwrite the row numbers, make sure the sheet is sorted properly!") == true) {
    var eventDetailsIterator = new EventDetailsIterator();
    var eventDetailsUpdater = new EventDetailsUpdater();
    eventDetailsIterator.iterate(eventDetailsUpdater);
  }
}

function onCheckCoordinator() {
  trace("onCheckCoordinator");
  var eventDetailsIterator = new EventDetailsIterator();
  var eventDetailsChecker = new EventDetailsChecker();
  eventDetailsIterator.iterate(eventDetailsChecker);
}


//=============================================================================================
// Class EventDetailsIterator
//
class EventDetailsIterator {
  constructor() {
    this.sourceRange = CellRange.getByName("EventDetails");
    let columnNamesRange = CellRange.getByName("EventDetailsColumnIds");
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
    trace(`EventRow.getCell --> ${Range.trace(cell)}`);
    return cell;
  }

  set(fieldName, value) { 
    let cell = this.getCell(fieldName);
    //trace("EventRow.set " + Range.trace(cell) + " = " + value);
    cell.setValue(value);
  }

  getA1Notation(fieldName) { 
    let columnNo = this.getColumnNo(fieldName);
    let cell = this.range.offset(0, columnNo, 1, 1);
    return cell.getA1Notation();
  }

  getSectionNo()      { return this.get("SectionNo"); }
  getItemNo()         { return this.get("ItemNo"); }
  isDecorTicked()     { return this.get("DecorTicked"); }
  isSupplierTicked()  { return this.get("SupplierTicked"); }
  isItineraryTicked() { return this.get("ItineraryTicked"); }
  getWho()            { return this.get("Who"); }
  getCategory()       { return this.get("Category"); }
  getStatus()         { return this.get("Status"); }
  get supplier()      { return this.get("Supplier"); }
  get isTitle()       { return this.getCategory() === "Title"; }  // Is this a title row?
  get title()         { return this.get("Description"); }
  getDate()           { return this.get("Date"); }
  getTime()           { return this.get("Time"); }
  getStartTime()      { return this.get("Time"); }
  getEndTime()        { return this.get("EndTime"); }
  getLocation()       { return this.get("Location"); }
  getDescription()    { return this.get("Description"); }
  getCurrency()       { return this.get("Currency"); }
  getCurrencySymbol() { return this.get("Currency") === GBP ? "£" : "€"; }
  getQuantity()       { return this.get("Quantity"); }
  getNativeUnitCost() { return this.get("NativeUnitCost");  }
  getMarkup()         { return this.get("Markup"); }
  getUnitPrice()      { return this.get("UnitPrice"); }
  getTotalPrice()     { return this.get("TotalPrice"); }
  getItemNotes()      { return this.get("ItemNotes"); }
  getNotes()          { return this.get("ItemNotes"); }
  getClientNotes()    { return ""; } // this.get("ItemNotes"); }
  getInventoryNotes() { return ""; } // this.get("ItemNotes"); }
  getLinks()          { return this.get("Links"); }
  
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

var EventDetailsUpdater = function() {
  trace("NEW " + this.trace());
}

EventDetailsUpdater.prototype.onBegin = function() {
  this.itemNo = 0;
  this.sectionNo = 0;
  this.eurGbpRate = Range.getByName("EURGBP").getValue();
  trace("EventDetailsUpdater.onBegin - EURGBP=" + this.eurGbpRate);
}
  
EventDetailsUpdater.prototype.onEnd = function() {
  trace("EventDetailsUpdater.onEnd - no-op");
}

EventDetailsUpdater.prototype.onTitle = function(row) {
  trace("EventDetailsUpdater.onTitle " + row.getTitle());
  this.itemNo = 0;
  ++this.sectionNo;
  if (row.getSectionNo() === "") { // Only set section id if empty
    row.set("SectionNo", this.generateSectionNo());
  }
  row.set("ItemNo", this.generateSectionNo());
}

EventDetailsUpdater.prototype.onRow = function(row) {
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
  if (row.getSectionNo() === "") { // Only set section id if empty
    row.set("SectionNo", this.generateSectionNo());
  }
  row.set("ItemNo", this.generateItemNo());
  row.set("UnitCost", Utilities.formatString('=IF(OR(%s="", %s="", %s=0), "", IF(%s="GBP", %s, %s / EURGBP))', currencyA1, nativeUnitCostA1, nativeUnitCostA1, currencyA1, nativeUnitCostA1, nativeUnitCostA1));
  row.set("TotalCost", Utilities.formatString('=IF(OR(%s="", %s=0, %s="", %s=0), "", %s * %s * (1-%s))', quantityA1, quantityA1, unitCostA1, unitCostA1, quantityA1, unitCostA1, commissionPercentageA1));
//  if (row.getMarkup() === "") { // Only set markup if empty
//    row.set("Markup", Utilities.formatString('=IF(OR(%s="", %s=0, %s="", %s=0), "", (%s-%s)/%s)', unitCostA1, unitCostA1, unitPriceA1, unitPriceA1, unitPriceA1, unitCostA1, unitCostA1));
//  }
  row.set("UnitPrice", Utilities.formatString('=IF(OR(%s="", %s=0), "", %s * ( 1 + %s))', unitCostA1, unitCostA1, unitCostA1, markupA1));
  row.set("TotalPrice", Utilities.formatString('=IF(OR(%s="", %s=0, %s="", %s=0), "", %s * %s)', quantityA1, quantityA1, unitPriceA1, unitPriceA1, quantityA1, unitPriceA1));
//  row.set("Commission", Utilities.formatString('=IF(OR(%s="", %s=0), "", %s * %s * %s)', commissionPercentageA1, commissionPercentageA1, quantityA1, unitCostA1, commissionPercentageA1));
  //if (=((hour(K215)*60+minute(K215))-(hour(J215)*60+minute(J215)))/60) 
}

EventDetailsUpdater.prototype.generateSectionNo = function() {
  return Utilities.formatString('#%02d', this.sectionNo);
}

EventDetailsUpdater.prototype.generateItemNo = function() {
  if (this.itemNo == 0) 
    return this.generateSectionNo();
  else
    return Utilities.formatString('%s-%02d', this.generateSectionNo(), this.itemNo);
}

EventDetailsUpdater.prototype.trace = function() {
  return "{EventDetailsUpdater}";
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
