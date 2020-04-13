function onUpdateExchangeRates() {
  trace("onUpdateExchangeRates");
  globalExchangeRates.updateRates();
  Dialog.notify("New Exchange Rate", 
                "From the European Central Bank\nhttps://exchangeratesapi.io/\n\nMid-rate (Marked up)\n" +
                `€1 = £${globalExchangeRates._eurToGbp} (${globalExchangeRates.eurToGbp})\n` + 
    `£1 = €${globalExchangeRates._gbpToEur.toPrecision(5)} (${globalExchangeRates.gbpToEur}`);
}


//=============================================================================================
// Class ExchangeRate

class ExchangeRate {
  constructor() {
    this.markup = 0.1;
    this.url = "https://api.exchangeratesapi.io/latest";
    this.rates = null;
    this._eurToGbp = 0.0;
    this._gbpToEur = 0.0;
  }

  updateRates() {
    let json = UrlFetchApp.fetch(ExchangeRate.url, {"muteHttpExceptions": true});
    //Dialog.notify("New Exchange Rates", json);
    this.rates = JSON.parse(json)["rates"];
    this._eurToGbp = this.rates["GBP"];
    this._gbpToEur = 1/this.eurToGbp;
    //Dialog.notify("New Exchange Rates", this.eurToGbp);
    trace("ExchangeRate.updateRates, MidRate: 1 EUR = " + this.eurToGbp + " GBP");
  }

  get eurToGbp() {
    return (this._eurToGbp/(1+this.markup)).toPrecision(2);
  }

  get gbpToEur() {
    return (this._gbpToEur*(1+this.markup)).toPrecision(3);
  }
}

var globalExchangeRates = new ExchangeRate();
