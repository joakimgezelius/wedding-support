// This script library is the main entry point for the Event Summary spreadsheets
// It can be accessed by referencing it using the script ID:
//
//  ScriptID: 1p3MGZxgnlKyi5cthDWd22lfSE6GWmy2LOhJrsSY05qjsk4DPs2n2xdEo
//  SDC key:  53741d44402f4b2c

globalLibName = "Event";

function onOpen() { // For backward compatibility
  trace("onOpen");
  addEventMenu();
}

function onEventSheetOpen(libName) {
  trace("onEventSheetOpen " + libName);
  globalLibName = libName;
  addEventMenu();
}

function onRotaSheetOpen(libName) {
  trace("onRotaSheetOpen");
  globalLibName = libName;
  addRotaMenu();
}

function onQuoteSheetOpen(libName) {
  trace("onQuoteSheetOpen");
  globalLibName = libName;
  addQuoteMenu();
}
