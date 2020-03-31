

function onUpdateExchangeRates() {
  trace('onUpdateExchangeRates');
  globalExchangeRates.updateRates();
  Dialog.notify('New Exchange Rate', 
                'From the European Central Bank\nhttps://exchangeratesapi.io/\n\nMid-rate (Marked up)\n' +
                '€1 = £' + globalExchangeRates.eurToGbp + ' (' + globalExchangeRates.getEurToGbp() + ')\n' + 
                '£1 = €' + globalExchangeRates.gbpToEur.toPrecision(5) + ' (' + globalExchangeRates.getGbpToEur() + ')');
}


//=============================================================================================
// Class ExchangeRate

var ExchangeRate = function() {
}

ExchangeRate.markup = 0.1;
ExchangeRate.url = 'https://api.exchangeratesapi.io/latest';

ExchangeRate.prototype.updateRates = function() {
  var json = UrlFetchApp.fetch(ExchangeRate.url, {'muteHttpExceptions': true});
//Dialog.notify("New Exchange Rates", json);
  this.rates = JSON.parse(json)['rates'];
  this.eurToGbp = this.rates['GBP'];
  this.gbpToEur = 1/this.eurToGbp;
//Dialog.notify("New Exchange Rates", this.eurToGbp);
  trace('ExchangeRate.updateRates, MidRate: 1 EUR = ' + this.eurToGbp + ' GBP');
}

ExchangeRate.prototype.getEurToGbp = function() {
  return (this.eurToGbp/(1+ExchangeRate.markup)).toPrecision(2);
}


ExchangeRate.prototype.getGbpToEur = function() {
  return (this.gbpToEur*(1+ExchangeRate.markup)).toPrecision(3);
}


var globalExchangeRates = new ExchangeRate();
