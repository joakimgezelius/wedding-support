//=============================================================================================
// Class EventDetails
//
const SortType = { time: "time", supplier: "supplier" };

class EventDetails {
  constructor() {
    this.range = Coordinator.eventDetailsRange;
    if (this.range) {
    }
    this.rowCount = this.range.height;
    this.values = this.range.values;     // NOTE: indexed from [0][0]
    trace("NEW " + this.trace);
  }

  // Method apply
  // Iterate over all event rows (using Range Row Iterator)
  //
  apply(handler) {
    trace(`${this.trace}.apply`);
    handler.onBegin();
    this.range.forEachRow((range) => {
      const row = new EventRow(range);
      if (row.isTitle) {
        handler.onTitle(row);
      } else {
        handler.onRow(row);
      }
    });
    handler.onEnd();
  }

  sort(type) {
    function compare(row1, row2) {
      let eventRow1 = new EventRow(this.range, null, row1); // Override the values, as we are sorting the values matrix only (risky)
      let eventRow2 = new EventRow(this.range, null, row2);
      return eventRow1.compare(eventRow2, type);
    }
    trace(`EventDetails.sort(${type}) ${this.trace}`);
    this.values.sort(compare);
  }
  
  get trace() {
    return `{EventDetails range=${this.range.trace} rowCount=${this.rowCount}`;
  }
}

//=============================================================================================
// Class EventRow
//
  
class EventRow extends RangeRow {
  
  constructor(range, rowOffset = null, values = null) {
    super(range, rowOffset, values);
  }

  get sectionNo()           { return this.get("ItemNo", "string").substr(0,3); }
  get itemNo()              { return this.get("ItemNo", "string"); }
  get isDecorTicked()       { return this.get("DecorTicked", "boolean"); }
  get isSupplierTicked()    { return this.get("SupplierTicked", "boolean"); }
  get isItineraryTicked()   { return this.get("ItineraryTicked", "boolean"); }
  get isTitle()             { return this.category.toLowerCase() === "title"; }    // Is this a title row?
  get isSubItem()           { return this.category.toLowerCase() === "part"; }     // Is this a sub-item?
  get isCancelled()         { return this.status.toLowerCase() === "cancelled"; }  // Is this item cancelled?
  get who()                 { return this.get("Who", "string"); }
  get category()            { return this.get("Category", "string"); }
  get status()              { return this.get("Status", "string"); }
  get supplier()            { return this.get("Supplier", "string"); }
  get title()               { return this.get("Description", "string"); }
  get date()                { return this.get("Date"); }
  get time()                { return this.get("Time"); }
  get startTime()           { return this.get("Time"); }
  get endTime()             { return this.get("EndTime"); }
  get location()            { return this.get("Location", "string"); }
  get description()         { return this.get("Description", "string"); }
  get currency()            { return this.get("Currency", "string").toUpperCase(); }
  get currencySymbol()      { return this.currency === "GBP" ? "£" : "€"; }
  get currencyFormat()      { return `${this.currencySymbol}#,##0`; }
  get quantity()            { return this.get("Quantity"); } // Accept blank
  get budgetUnitCost()      { return this.get("BudgetUnitCost");  } // Accept blank
  get nativeUnitCost()      { return this.get("NativeUnitCost");  } // Accept blank
  get markup()              { return this.get("Markup"); } // Accept ref errors
  get commissionPercentage(){ return this.get("CommissionPercentage"); } // Accept ref errors
  get unitPrice()           { return this.get("UnitPrice"); } // Accept blank
  get totalPrice()          { return this.get("TotalPrice", "number"); }
  get itemNotes()           { return this.get("ItemNotes", "string"); }
  get notes()               { return this.get("ItemNotes", "string"); }
  get clientNotes()         { return ""; } // this.get("ItemNotes", "string"); }
  get inventoryNotes()      { return ""; } // this.get("ItemNotes", "string"); }
  get links()               { return this.get("Links", "string"); }

  set itemNo(value)         { this.set("ItemNo", value); }
  set category(value)       { this.set("Category", value); }
  set supplier(value)       { this.set("Supplier", value); }
  set description(value)    { this.set("Description", value); }
  set currency(value)       { this.set("Currency", value); }
  set markup(value)         { this.set("Markup", value); }
  set nativeUnitCost(value) { this.set("NativeUnitCost", value); }
  set unitCost(value)       { this.set("UnitCost", value); }
  set totalCost(value)      { this.set("TotalCost", value); }
  set unitPrice(value)      { this.set("UnitPrice", value); }
  set totalPrice(value)     { this.set("TotalPrice", value); }
  set commissionPercentage(value) { this.set("CommissionPercentage", value); }
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
