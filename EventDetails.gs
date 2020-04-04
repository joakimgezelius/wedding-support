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
var EventDetailsIterator = function() {
  this.range = Range.getByName("EventDetails");
  var columnNamesRange = Range.getByName("EventDetailsColumnIds");
  this.rowCount = this.range.getHeight();
  this.data = this.range.getValues();
//  var columnNamesRange = this.range.offset(-1, 0, 1); // We expect to find the column names in the row above the data
  EventRow.columnNames = columnNamesRange.getValues()[0];
  trace("NEW " + this.trace());
  trace("Columns: " + EventRow.columnNames);
  for (var column = 0; column < EventRow.columnNames.length; ++column) {
    var columnName = EventRow.columnNames[column];
    if (columnName !== "") {
      EventRow[columnName] = column;
//    trace("Column: " + columnName);
    }
  }
}

// Method iterate
// Iterate over all event rows
//
EventDetailsIterator.prototype.iterate = function(handler) {
  trace("EventDetailsIterator.iterate " + this.trace());
  handler.onBegin();
  for (var rowOffset = 0; rowOffset < this.rowCount; rowOffset++) {
    var rowRange = this.range.offset(rowOffset, 0, 1);
    var row = new EventRow(this.data[rowOffset], rowOffset, rowRange);
    if (row.isTitle()) {
      handler.onTitle(row);
    } else {
      handler.onRow(row);
    }
  }  
  handler.onEnd();
}
  
EventDetailsIterator.prototype.sortByTime = function() {
  compareTime = function(row1, row2) {
    eventRow1 = new EventRow(row1);
    eventRow2 = new EventRow(row2);
    return eventRow1.compareTime(eventRow2);
  }
  trace("EventDetailsIterator.sortByTime " + this.trace());
  this.data.sort(compareTime);
}
  
EventDetailsIterator.prototype.trace = function() {
  return "{EventDetailsIterator range=" + Range.trace(this.range) + ", rowCount=" + this.rowCount + "}";
}
  

//=============================================================================================
// Class EventRow
//
var EventRow = function(data, offset, range) {
  this.data = data;
  this.offset = offset;
  this.range = range;
}

EventRow.prototype.get = function(fieldName) { 
  if (!EventRow.hasOwnProperty(fieldName)) Error.fatal("Unknown EventRow column: " + fieldName);
  var columnNo = EventRow[fieldName];
  return this.data[columnNo];
}

EventRow.prototype.getCell = function(fieldName) { 
  if (!EventRow.hasOwnProperty(fieldName)) Error.fatal("Unknown EventRow column: " + fieldName);
  var columnNo = EventRow[fieldName];
  var cell = this.range.offset(0, columnNo, 1, 1);
  trace("EventRow.getCell --> " + Range.trace(cell));
  return cell;
}

EventRow.prototype.set = function(fieldName, value) { 
  var cell = this.getCell(fieldName);
//trace("EventRow.set " + Range.trace(cell) + " = " + value);
  cell.setValue(value);
}

EventRow.prototype.getA1Notation = function(fieldName) { 
  if (!EventRow.hasOwnProperty(fieldName)) Error.fatal("Unknown EventRow column: " + fieldName);
  var columnNo = EventRow[fieldName];
  var cell = this.range.offset(0, columnNo, 1, 1);
  return cell.getA1Notation();
}

EventRow.prototype.getSectionNo       = function() { return this.get("SectionNo"); }
EventRow.prototype.getItemNo          = function() { return this.get("ItemNo"); }
EventRow.prototype.isDecorTicked      = function() { return this.get("DecorTicked"); }
EventRow.prototype.isSupplierTicked   = function() { return this.get("SupplierTicked"); }
EventRow.prototype.isItineraryTicked  = function() { return this.get("ItineraryTicked"); }
EventRow.prototype.getWho             = function() { return this.get("Who"); }
EventRow.prototype.getCategory        = function() { return this.get("Category"); }
EventRow.prototype.getStatus          = function() { return this.get("Status"); }
EventRow.prototype.getSupplier        = function() { return this.get("Supplier"); }
EventRow.prototype.isTitle            = function() { return this.getCategory() === "Title"; }  // Is this a title row?
EventRow.prototype.getTitle           = function() { return this.get("Description"); }
EventRow.prototype.getDate            = function() { return this.get("Date"); }
EventRow.prototype.getTime            = function() { return this.get("Time"); }
EventRow.prototype.getStartTime       = function() { return this.get("Time"); }
EventRow.prototype.getEndTime         = function() { return this.get("EndTime"); }
EventRow.prototype.getLocation        = function() { return this.get("Location"); }
EventRow.prototype.getDescription     = function() { return this.get("Description"); }
EventRow.prototype.getCurrency        = function() { return this.get("Currency"); }
EventRow.prototype.getCurrencySymbol  = function() { return this.get("Currency") === GBP ? "£" : "€"; }
EventRow.prototype.getQuantity        = function() { return this.get("Quantity"); }
EventRow.prototype.getNativeUnitCost  = function() { return this.get("NativeUnitCost");  }
EventRow.prototype.getMarkup          = function() { return this.get("Markup"); }
EventRow.prototype.getUnitPrice       = function() { return this.get("UnitPrice"); }
EventRow.prototype.getTotalPrice      = function() { return this.get("TotalPrice"); }
EventRow.prototype.getItemNotes       = function() { return this.get("ItemNotes"); }
EventRow.prototype.getNotes           = function() { return this.get("ItemNotes"); }
EventRow.prototype.getClientNotes     = function() { return ""; } // this.get("ItemNotes"); }
EventRow.prototype.getInventoryNotes  = function() { return ""; } // this.get("ItemNotes"); }
EventRow.prototype.getLinks           = function() { return this.get("Links"); }
  
EventRow.prototype.compareTime = function(other) {
  var result = 0;
  if (this.getDate() < other.getDate()) result = -1;
  else if (this.getDate() > other.getDate()) result = 1;
  // Same date, now compare times
  else if (this.getTime() < other.getTime()) result = -1;
  else if (this.getTime() > other.getTime()) return 1;
  else result = 0; // Both date & time are the same
  trace("EventRow.compareTime " + result);
  return result;
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

var EventDetailsChecker = function() {
  trace("NEW " + this.trace());
}

EventDetailsChecker.prototype.onBegin = function() {
  trace("EventDetailsChecker.onBegin - no-op");
}
  
EventDetailsChecker.prototype.onEnd = function() {
  trace("EventDetailsChecker.onEnd - no-op");
}

EventDetailsChecker.prototype.onTitle = function(row) {
  trace("EventDetailsChecker.onTitle " + row.getTitle());
}

EventDetailsChecker.prototype.onRow = function(row) {
  trace("EventDetailsChecker.onRow ");
}
  
EventDetailsChecker.prototype.trace = function() {
  return "{EventDetailsChecker}";
}
