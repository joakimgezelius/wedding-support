// This script library is the main entry point for the Wedding Summary spreadsheets
// It can be accessed by referencing it using the script ID:
//
//  ScriptID: 1CiKOqQNFxdyAR5PZ7GaJDnGyZfrh-b-rJIlyCFAI2ABvu8yV1pCmm3ER
//
//  ScriptID (Jithin): 1mXJiRin063xGwitgWu-0-rUepMTe_GTO1A-oJ5vJnTtIlZLZPnVhpd76
//  ScriptID (Shaikh): 1uWdnau49TWbINbe_I4eY_JJDswE_J2hQiEB-iKu35-DpdIJf0-EdOn6P

globalLibMenuTag = ""; // Set to unique value for each library, e.g. " (Shaikh)"
globalLibId = "Prod";

function onWeddingSheetOpen(libName) {
  trace(`onWeddingSheetOpen libName=${libName}`);
  Menu.addEventMenu(libName);
  Menu.addAsanaMenu(libName);
}

function onRotaSheetOpen(libName) {
  trace(`onRotaSheetOpen libName=${libName}`);
  Menu.addRotaMenu(libName);
}

function onEventCoordinationSheetOpen(libName) {
  trace(`onEventCoordinationSheetOpen libName=${libName}`);
  Menu.addEventCoordinationMenu(libName);
}

function onDecorSheetOpen(libName) {
  trace(`onDecorSheetOpen libName=${libName}`);
  Menu.addDecorPriceListMenu(libName);
}

function onShopSalesSheetOpen(libName) {
  trace(`onShopSalesSheetOpen libName=${libName}`);
  Menu.addShopSalesListMenu(libName);
}

function onQuoteSheetOpen(libName) {
  trace(`onQuoteSheetOpen libName=${libName}`);
  Menu.addQuoteMenu(libName);
}

function onEnqiriesSheetOpen(libName) {
  trace(`onEnqiriesSheetOpen libName=${libName}`);
  Menu.addEnqiriesMenu(libName);
}

function onWeddingPackagesSheetOpen(libName) {
  trace(`onWeddingPackagesSheetOpen libName=${libName}`);
  Menu.addWeddingPackagesMenu(libName);
}

function onProductsSheetOpen(libName) {
  trace(`onProductsSheetOpen libName=${libName}`);
  Menu.addProductsSheetMenu(libName);
}

function onUtilitiesSheetOpen(libName) {
  trace(`onUtilitiesSheetOpen libName=${libName} LibId=${globalLibId}`);
  Menu.addUtilitiesSheetMenu(libName);
}

function onTestMenu(libName) {
  trace(`onTestOpen libName=${libName}`);
  Menu.addTestMenu(libName);
}
