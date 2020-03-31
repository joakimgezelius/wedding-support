
function onUpdateQuote() {
  trace("onUpdateQuote");
  Dialog.confirm("Update Quote", "this will update a quote!");
  Quote.update();
}


function onCreateEstimateSummary() {
  trace("onCreateEstimateSummary");
  Dialog.confirm("Create Estimate Summary", "this will create an estimate summary!");
}


var Quote = function() {
}

Quote.priceList = null;

Quote.init = function() {
  trace("Quote.init");
  if (this.priceList == null) {
    this.priceList = new PriceList();
  }
}

Quote.update = function() {
  trace("Quote.update");
  this.init();
  this.priceList.update();
}

Quote.onButton = function() { // Not prototype in order to be callable statically
  Dialog.notify(globalLibName + ".Quote.onButton", "do it!");
}

Quote.onEdit = function(event) {
  Dialog.notify("Quote.onEdit", event.range.getA1Notation());
}
