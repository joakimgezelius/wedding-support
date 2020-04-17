//=============================================================================================
// Class EventDetailsIterator
//
const SortType = { time: "time", supplier: "supplier" };

class EventDetailsIterator {
  constructor() {
    this.sourceRange = Range.getByName("EventDetails");
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
                      
  sort(type) {
    function compare(row1, row2) {
      let eventRow1 = new EventRow(row1);
      let eventRow2 = new EventRow(row2);
      return eventRow1.compare(eventRow2, type);
    }
    trace(`EventDetailsIterator.sort(${type}) ${this.trace}`);
    this.data.sort(compare);
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
//  trace(`EventRow.getCell ${columnName} --> ${Range.trace(cell)}`);
    return cell;
  }

  set(columnName, value) { 
    let cell = this.getCell(columnName);
    trace(`EventRow.set ${columnName} ${Range.trace(cell)}  = ${value}`);
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
  get isTitle()           { return this.category.toUpperCase() === "TITLE"; }  // Is this a title row?
  get isSubItem()         { return this.category.toUpperCase() === "PART"; }   // Is this a sub-item?
  get who()               { return this.get("Who"); }
  get category()          { return this.get("Category"); }
  get status()            { return this.get("Status"); }
  get supplier()          { return this.get("Supplier"); }
  get title()             { return this.get("Description"); }
  get date()              { return this.get("Date"); }
  get time()              { return this.get("Time"); }
  get startTime()         { return this.get("Time"); }
  get endTime()           { return this.get("EndTime"); }
  get location()          { return this.get("Location"); }
  get description()       { return this.get("Description"); }
  get currency()          { return this.get("Currency").toUpperCase(); }
  get currencySymbol()    { return this.currency === "GBP" ? "£" : "€"; }
  get currencyFormat()    { return `{this.currencySymbol}#,##0`; }
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

  compare(other, type) { // To support sorting of rows
    let result = 0;
    if (type === SortType.supplier) { // complare suppliers
      result = this.compareSupplier(other);
    }
    if (result !== 0) return result;
    // supplier is the same, now compare dates
    result = this.compareDate(other);
    if (result !== 0) return result;
    // Same date, now compare times
    result = this.compareTime(other);
    return result;
  }

  compareSupplier(other) {
    let myValue = this.supplier;
    let otherValue = other.supplier;
    return myValue < otherValue ? -1 : (myValue > otherValue ? 1 : 0);
  }

  compareDate(other) {
    let myValue = this.date;
    let otherValue = other.date;
    return myValue < otherValue ? -1 : (myValue > otherValue ? 1 : 0);
  }

  compareTime(other) {
    let myValue = this.time;
    let otherValue = other.time;
    return myValue < otherValue ? -1 : (myValue > otherValue ? 1 : 0);
  }

}
