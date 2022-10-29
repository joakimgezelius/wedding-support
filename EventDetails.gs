//=============================================================================================
// Class EventDetails
//
const SortType = { date: "date", time: "time", supplier: "supplier", supplier_location: "supplier_location", staff: "staff" };

class EventDetails {
  
  constructor(rangeName = "EventDetails") { // NOTE: EventDetails is the default range name in the client sheet coordinator, we however also use this class to manage price lists.
    trace("constructing EventDetails object...");
    this.range = Range.getByName(rangeName).loadColumnNames();
    this.rowCount = this.range.height;
    //this.values = this.range.values;     // NOTE: indexed from [0][0]
    trace("NEW " + this.trace);
  }

  // Method apply
  // Iterate over all event rows (using Range Row Iterator), call handler methods 
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

  // NOTE: Sorting the underlying array of values for the range is dangerous, as the range itself isn't sorted!
  //
  sort(type) {
    trace(`EventDetails.sort(${type}) ${this.trace}`);
    this.range.values.sort((row1, row2) => {
      let eventRow1 = new EventRow(this.range, row1); // Override the values, as we are sorting the values matrix only (risky!)
      let eventRow2 = new EventRow(this.range, row2);
      return eventRow1.compare(eventRow2, type);
    });
  }
  
  get trace() {
    return `{EventDetails range=${this.range.trace} rowCount=${this.rowCount}`;
  }
    
}

//=============================================================================================
// Class EventRow
//
  
class EventRow extends RangeRow {
  
  constructor(range, values = null) {
    super(range, values);
  }

  get sectionNo()           { return this.get("ItemNo", "string").substr(0,3); }
  get itemNo()              { return this.get("ItemNo", "string"); }
  get isDecorTicked()       { return this.get("DecorTicked", "boolean"); }
  get isSupplierTicked()    { return this.get("VenueTicked", "boolean"); }
  get isStoreTicked()       { return this.get("StoreTicked"); }      // To pick store items through query & accept it blank
  //get isStaffTicked()     { return this.get("StaffTicked", "boolean"); }
  get isItineraryTicked()   { return this.get("ItineraryTicked", "boolean"); }
  get isTitle()             { return this.category.toLowerCase() === "title"; }    // Is this a title row?
  get isSubItem()           { return this.category.toLowerCase() === "part"; }     // Is this a sub-item?
  get isInStock()           { return this.category.toLowerCase() === "in stock"; }   // Is this item in stock?
  get isCancelled()         { return this.status.toLowerCase() === "cancelled"; }  // Is this item cancelled?
  get who()                 { return this.get("Who", "string"); }
  get incharge()            { return this.get("InCharge", "string"); }
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
  get quantity()            { return this.get("Quantity"); } // Accept blank
  get currency()            { return this.get("Currency", "string").toUpperCase(); }
  get currencySymbol()      { return this.currency === "GBP" ? "£" : "€"; }
  get currencyFormat()      { return `${this.currencySymbol}#,##0.00`; }
  get budgetUnitCost()      { return this.get("BudgetUnitCost");  } // Accept blank
  get nativeUnitCost()      { return this.get("NativeUnitCost");  } // Accept blank
  get nativeUnitCostWithVAT() { return this.get("NativeUnitCostWithVAT"); } // Accept blank
  get unitCost()            { return this.get("UnitCost");  } // Accept blank
  get totalNativeGrossCost(){ return this.quantity * this.nativeUnitCostWithVAT; }
  get totalGrossCost()      { return this.get("TotalGrossCost");  } // Accept blank
  get totalNettCost()       { return this.get("TotalCost");  } // Accept blank
  get markup()              { return this.get("Markup"); } // Accept ref errors
  get commissionPercentage(){ return this.get("CommissionPercentage"); } // Accept ref errors
  get unitPrice()           { return this.get("UnitPrice"); } // Accept blank
  get totalPrice()          { return this.get("TotalPrice", "number"); }
  get itemNotes()           { return this.get("ItemNotes", "string"); }
  get notes()               { return this.get("ItemNotes", "string"); }
  get clientNotes()         { return ""; } // this.get("ItemNotes", "string"); }
  get inventoryNotes()      { return ""; } // this.get("ItemNotes", "string"); }
  get paymentMethod()       { return this.get("PaymentMethod", "string"); }
  get paymentStatus()       { return this.get("PaymentStatus", "string"); }
  get links()               { return this.get("Links", "string"); }
  get isSelected()          { return this.get("Selected"); } // Accept blank
 
  set itemNo(value)         { this.set("ItemNo", value); }
  set isStoreTicked(value)  { this.set("StoreTicked", value); }   // To inject query to tick true/false based on other column values
  set category(value)       { this.set("Category", value); }
  set status(value)         { this.get("Status", value); }
  set supplier(value)       { this.set("Supplier", value); }
  set description(value)    { this.set("Description", value); }
  set quantity(value)       { this.set("Quantity", value); }
  set currency(value)       { this.set("Currency", value); }
  set markup(value)         { this.set("Markup", value); }
  set nativeUnitCost(value) { this.set("NativeUnitCost", value); }
  set nativeUnitCostWithVAT(value) { this.set("NativeUnitCostWithVAT", value); }
  set unitCost(value)       { this.set("UnitCost", value); }
  set totalGrossCost(value) { this.set("TotalGrossCost", value); }
  set totalNettCost(value)  { this.set("TotalCost", value); }
  set unitPrice(value)      { this.set("UnitPrice", value); }
  set totalPrice(value)     { this.set("TotalPrice", value); }
  set commissionPercentage(value) { this.set("CommissionPercentage", value); }
  set commission(value)     { this.set("Commission", value); }
  set isSelected(value)     { this.set("Selected", value); }

  compare(other, type) { // To support sorting of rows
    let result = 0;
    if (type === SortType.supplier) { // compare suppliers
      result = this.compareSupplier(other);
    }
    else if (type === SortType.supplier_location) { // compare suppliers, then location
      result = this.compareSupplierLocation(other);
    }
    else if (type === SortType.staff) { // compare staff
      result = this.compareStaff(other);
    }
    if (result !== 0) return result;
    // supplier or staff is the same, now compare dates
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

  compareSupplierLocation(other) {
    let myValue = this.supplier;
    let otherValue = other.supplier;
    return myValue < otherValue ? -1 : (myValue > otherValue ? 1 : this.compareLocation(other));
  }

  compareLocation(other) {
    let myValue = this.location;
    let otherValue = other.location;
    return myValue < otherValue ? -1 : (myValue > otherValue ? 1 : 0);
  }

  compareStaff(other) {
    let myValue = this.who;
    let otherValue = other.who;
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