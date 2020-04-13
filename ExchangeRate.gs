function onUpdateExchangeRates() {
  trace("onUpdateExchangeRates");
  globalExchangeRates.updateRates();
  Dialog.notify("New Exchange Rate", 
                "From the European Central Bank\nhttps://exchangeratesapi.io/\n\nMid-rate (Marked up)\n" +
                `€1 = £${globalExchangeRates.myEurToGbp} (${globalExchangeRates.eurToGbp})\n` + 
    `£1 = €${globalExchangeRates.gbpToEur.toPrecision(5)} (${globalExchangeRates.gbpToEur}`);
}


//=============================================================================================
// Class ExchangeRate

class ExchangeRate {
  constructor() {
    this.markup = 0.1;
    this.url = "https://api.exchangeratesapi.io/latest";
    this.rates = null;
    this.myEurToGbp = 0.0;
    this.myGbpToEur = 0.0;
  }

  updateRates() {
    let json = UrlFetchApp.fetch(ExchangeRate.url, {"muteHttpExceptions": true});
    //Dialog.notify("New Exchange Rates", json);
    this.rates = JSON.parse(json)["rates"];
    this.myEurToGbp = this.rates["GBP"];
    this.myGbpToEur = 1/this.eurToGbp;
    //Dialog.notify("New Exchange Rates", this.eurToGbp);
    trace("ExchangeRate.updateRates, MidRate: 1 EUR = " + this.eurToGbp + " GBP");
  }

  get eurToGbp() {
    return (this.eurToGbp/(1+this.markup)).toPrecision(2);
  }

  get gbpToEur() {
    return (this.gbpToEur*(1+this.markup)).toPrecision(3);
  }
}

var globalExchangeRates = new ExchangeRate();
