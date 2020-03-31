var PriceList = function() {
  this.range = Range.getByName("PriceList");
  this.rowCount = this.range.getHeight();
  this.data = this.range.getValues();
  trace("NEW " + this.trace());
}

PriceList.prototype.update = function() {
  this.gatherCategories();
}

PriceList.prototype.gatherCategories = function() {
  trace("PriceList.gatherCategories");
  for (var rowOffset = 0; rowOffset < this.rowCount; rowOffset++) {
    var rowRange = this.range.offset(rowOffset, 0, 1);
  }
}

PriceList.prototype.trace = function() {
  return "{PriceList " + Range.trace(this.range) + "}";
}
