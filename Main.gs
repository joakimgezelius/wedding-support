// This script library is the main entry point for the Wedding Summary spreadsheets
// It can be accessed by referencing it using the script ID:
//
//  ScriptID: 1CiKOqQNFxdyAR5PZ7GaJDnGyZfrh-b-rJIlyCFAI2ABvu8yV1pCmm3ER
//
//  ScriptID (Jithin): 1mXJiRin063xGwitgWu-0-rUepMTe_GTO1A-oJ5vJnTtIlZLZPnVhpd76
//  ScriptID (Shaikh): 1uWdnau49TWbINbe_I4eY_JJDswE_J2hQiEB-iKu35-DpdIJf0-EdOn6P

globalLibName = "Undefined";

function onOpen() { // For backward compatibility
  trace(`onOpen globalLibName=${globalLibName}`);
  addWeddingMenu();
}

function onWeddingSheetOpen(libName) { // onWeddingSheetOpen("WedLib");
  trace(`onWeddingSheetOpen globalLibName=${libName}`);
  globalLibName = libName;
  addWeddingMenu();
}

function onRotaSheetOpen(libName) {
  trace(`onRotaSheetOpen globalLibName=${libName}`);
  globalLibName = libName;
  addRotaMenu(); // in Menu.gs
}

function onQuoteSheetOpen(libName) {
  trace(`onQuoteSheetOpen globalLibName=${libName}`);
  globalLibName = libName;
  addQuoteMenu();
}

function onEnqiriesSheetOpen(libName) {
  trace(`onEnqiriesSheetOpen globalLibName=${libName}`);
  globalLibName = libName;
  addEnqiriesMenu();
}

function onWeddingPackagesSheetOpen(libName) {
  trace(`onWeddingPackagesSheetOpen globalLibName=${libName}`);
  globalLibName = libName;
  addWeddingPackagesMenu();
}

function onProductsSheetOpen(libName) {
  trace(`onProductsSheetOpen globalLibName=${libName}`);
  globalLibName = libName;
  addProductsSheetMenu();
  addWeddingMenu();
  //
}
