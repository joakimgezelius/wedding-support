// This script library is the main entry point for the Wedding Summary spreadsheets
// It can be accessed by referencing it using the script ID:
//
//  ScriptID: 1CiKOqQNFxdyAR5PZ7GaJDnGyZfrh-b-rJIlyCFAI2ABvu8yV1pCmm3ER
//  SDC key:  ef0d427283541f50

globalLibName = "Undefined";

function onOpen() { // For backward compatibility
  trace("onOpen globalLibName=" +globalLibName);
  addWeddingMenu();
}

function onWeddingSheetOpen(libName) { // onWeddingSheetOpen("WedLib");
  trace("onWeddingSheetOpen globalLibName=" + libName);
  globalLibName = libName;
  addWeddingMenu();
}

function onRotaSheetOpen(libName) {
  trace("onRotaSheetOpen globalLibName=" + libName);
  globalLibName = libName;
  addRotaMenu();
}

function onQuoteSheetOpen(libName) {
  trace("onQuoteSheetOpen globalLibName=" + libName);
  globalLibName = libName;
  addQuoteMenu();
}
