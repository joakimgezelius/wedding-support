//=============================================================================================
// Class DecorPriceList
//

class DecorPriceList {

  constructor() {
    trace("constructing Decor Price List object...");
    this.range = DecorPriceListBuilder.decorPriceListRange;
    this.rowCount = this.range.height;
    trace("NEW " + this.trace);
  }

  // Method apply
  // Iterate over all rows (using Range Row Iterator), call handler methods 

  apply(handler) {
    trace(`${this.trace}.apply`);
    this.range.forEachRow((range) => {
      const row = new DecorPriceListRow(range);
      handler.onRow(row);
    });
    handler.onEnd();
  }

  get trace() {
    return `{Decor Price List range=${this.range.trace} rowCount=${this.rowCount}`;
  }
}


//=============================================================================================
// Class PriceListRow
//

class DecorPriceListRow extends RangeRow {

  constructor(range, values = null) {
    super(range, values);
  }

  get currency()            { return this.get("Currency", "string").toUpperCase(); }
  get currencySymbol()      { return this.currency === "GBP" ? "£" : "€"; }
  get currencyFormat()      { return `${this.currencySymbol}#,##0.00`; }
  get nativeUnitCost()      { return this.get("NativeUnitCost");  } // Accept blank
  get nativeUnitCostWithVAT() { return this.get("NativeUnitCostWithVAT"); } // Accept blank
  get unitCost()            { return this.get("UnitCost");  } // Accept blank
  get markup()              { return this.get("Markup"); } // Accept ref errors
  get commissionPercentage(){ return this.get("CommissionPercentage"); } // Accept ref errors
  get unitPrice()           { return this.get("UnitPrice"); } // Accept blank
  
  set currency(value)       { this.set("Currency", value); }  
  set nativeUnitCost(value) { this.set("NativeUnitCost", value); }
  set nativeUnitCostWithVAT(value) { this.set("NativeUnitCostWithVAT", value); }
  set unitCost(value)       { this.set("UnitCost", value); }
  set markup(value)         { this.set("Markup", value); }
  set commissionPercentage(value) { this.set("CommissionPercentage", value); }
  set unitPrice(value)      { this.set("UnitPrice", value); }

}
