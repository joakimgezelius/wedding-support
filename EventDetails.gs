//=============================================================================================
// Class EventDetailsIterator
//
class EventDetailsIterator {
  constructor() {
    this.sourceRange = CRange.getByName("EventDetails");
    this.rowCount = this.sourceRange.height;
    this.data = this.sourceRange.values; // NOTE: indexed from [0][0]
    trace("NEW " + this.trace);
  }

  // Method iterate
  // Iterate over all event rows
  //
  iterate(handler) {
    trace("EventDetailsIterator.iterate " + this.trace);
    handler.onBegin();
    for (var rowOffset = 0; rowOffset < this.rowCount; rowOffset++) {
      let rowRange = this.sourceRange.range.offset(rowOffset, 0, 1);
      let row = new EventRow(this.data[rowOffset], rowOffset, rowRange);
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
    EventRow.init();
  }

  static init() { // Static initialisation
    if (!("columnNumbers" in EventRow)) {
      EventRow.columnNumbers = new NamedColumns("EventRow", "EventDetailsColumnIds");
    }
  }
  
  get(columnName) {
    return this.data[EventRow.columnNumbers.getColumnNumber(columnName)];
  }

  getCell(columnName) { 
    let columnNumber = EventRow.columnNumbers.getColumnNumber(columnName);
    let cell = this.range.offset(0, columnNumber, 1, 1);
//  trace(`EventRow.getCell ${columnName} --> ${CRange.trace(cell)}`);
    return cell;
  }

  set(columnName, value) { 
    let cell = this.getCell(columnName);
    trace(`EventRow.set ${columnName} ${CRange.trace(cell)}  = ${value}`);
    cell.setValue(value);
  }

  getA1Notation(columnName) { 
    let cell = this.getCell(columnName);
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
  get budgetUnitCost()    { return this.get("BudgetUnitCost");  }
  get nativeUnitCost()    { return this.get("NativeUnitCost");  }
  get markup()            { return this.get("Markup"); }
  get unitPrice()         { return this.get("UnitPrice"); }
  get totalPrice()        { return this.get("TotalPrice"); }
  get itemNotes()         { return this.get("ItemNotes"); }
  get notes()             { return this.get("ItemNotes"); }
  get clientNotes()       { return ""; } // this.get("ItemNotes"); }
  get inventoryNotes()    { return ""; } // this.get("ItemNotes"); }
  get links()             { return this.get("Links"); }

  set itemNo(value)         { this.set("ItemNo", value); }
  set markup(value)         { this.set("Markup", value); }
  set nativeUnitCost(value) { this.set("NativeUnitCost", value); }
  set unitCost(value)       { this.set("UnitCost", value); }
  set totalCost(value)      { this.set("TotalCost", value); }
  set unitPrice(value)      { this.set("UnitPrice", value); }
  set totalPrice(value)     { this.set("TotalPrice", value); }
  set commission(value)     { this.set("Commission", value); }

  compareTime(other) { // To support sorting of rows
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


