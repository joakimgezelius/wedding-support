function addWeddingMenu() {
  trace("> Adding custom wedding menu");
  let ui = SpreadsheetApp.getUi();
  ui.createMenu("Event" + globalLibMenuTag)      
      //.addItem("Clear Sheet", globalLibName + ".onClearSheet")
      //.addSeparator()
      //.addItem("Pull Client Information", globalLibName + ".onPullClientInformation") // in client.gs
      //.addSeparator()
      .addItem("Update Coordinator (no over-writes)", globalLibName + ".onUpdateCoordinator")
      .addItem("Force-Update Coordinator - over-writes data", globalLibName + ".onUpdateCoordinatorForced")
      //.addItem("Check Coordinator", globalLibName + ".onCheckCoordinator")
      .addItem("Update Budget", globalLibName + ".onUpdateBudget")
      .addItem("Update Client Itinerary", globalLibName + ".onUpdateClientItinerary") // In itinerary.gs
      .addItem("Update Decor Sumamry", globalLibName + ".onUpdateDecorSummary")
      .addItem("Update Account Summary", globalLibName + ".onUpdateAccountSummary")
      .addItem("Update Supplier Costing", globalLibName + ".onUpdateSupplierCosting")
      .addItem("Update Supplier Itinerary", globalLibName + ".onUpdateSupplierItinerary")
      //.addItem("Update Staff Itinerary", globalLibName + ".onUpdateStaffItinerary")
      //.addItem("Update Rota", globalLibName + ".onUpdateRota")
      .addSeparator()
      .addItem("Update Exchange Rates", globalLibName + ".onUpdateExchangeRates")
      .addSeparator()
      .addSubMenu(ui.createMenu("Prices")
                  .addItem("Import price list", globalLibName + ".onImportPriceList")
                  .addItem("tbd...", globalLibName + ".onImportPriceList")
                  )
      .addSeparator()
      //.addItem("Create New Client Sheet", globalLibName + ".onCreateNewClientSheet")
      .addItem("Apply Format Template", globalLibName + ".onApplyFormat")    // In SheetFormat.gs
      .addItem("Format Coordinator", globalLibName + ".onFormatCoordinator")
      .addSeparator()
      .addSubMenu(ui.createMenu("Email")
                  .addItem("Prepare client email draft 1", globalLibName + ".onCreateFirstEmailDraft")
                  .addItem("Send client email 1", globalLibName + ".onSendFirstEmail")
                  )
//    .addSubMenu(ui.createMenu("Tasks")
//                  .addItem("Small Wedding", globalLibName + ".onShowSmallWeddingTasks")
//                  .addItem("Big Wedding", globalLibName + ".onShowBigWeddingTasks")
//                  .addItem("Show All", globalLibName + ".onShowAllTasks")
//                  )
      .addSeparator()
      .addSubMenu(ui.createMenu("Helpers")
                  .addItem("Reverse Mark-up Calculations", globalLibName + ".onReverseMarkupCalculations")
                  .addItem("Set Colour", globalLibName + ".onSetColour")
                  .addItem("Test Case 1", globalLibName + ".onTestCase1")
                  .addItem("Test Case 2", globalLibName + ".onTestCase2")
                  )
      .addToUi();
  trace("< Custom wedding menu added");
}

function addTestMenu() {
  trace("> Adding custom test menu");
  let ui = SpreadsheetApp.getUi();
  ui.createMenu("Test" + globalLibMenuTag)
    .addItem("Update Dynamic Itinerary", "Event.onUpdateDynamicItinerary")
/*
    .addSubMenu(ui.createMenu("Testing")
      .addItem("Clear", globalLibName + ".onClearCoordinatorSummary")
      .addItem("List All Sheets", globalLibName + ".onListAllSheets")
      .addItem("List Coordinator Sheets", globalLibName + ".onListCoordinatorSheets")
      .addItem("Test Coordinator Summary", globalLibName + ".onUpdateCoordinatorSummary")
      .addItem("Show Details", globalLibName + ".onShowDetails")
      .addItem("Hide Details", globalLibName + ".onHideDetails")
      .addSeparator()
    )
    .addSubMenu(ui.createMenu("Log")
      .addItem("View", globalLibName + ".onLogView")
      .addItem("Clear", globalLibName + ".onLogClear")
    )
    */
    .addToUi();
  trace("< Custom  test menu added");
}

function addRotaMenu() {
  trace("> Adding custom rota menu");
  let ui = SpreadsheetApp.getUi();
  ui.createMenu("Rota" + globalLibMenuTag)
      .addItem("Update Rota", globalLibName + ".onUpdateRota")
      .addItem("Update Transportation", globalLibName + ".onUpdateTransportation")
      .addItem("Update Things-to-Buy", globalLibName + ".onUpdateThingsToBuy")
      .addItem("Update Things-in-Store", globalLibName + ".onUpdateThingsInStore")
      .addItem("Update Consumable", globalLibName + ".onUpdateConsumable")
      .addItem("Update In Shop", globalLibName + ".onUpdateInShop")
      .addItem("Update Service", globalLibName + ".onUpdateService")
      .addSeparator()
      .addItem("Activity Colour Coding", globalLibName + ".onActivityColouring")
      .addItem("Supplier Colour Coding", globalLibName + ".onSupplierColouring")
      .addItem("Location Colour Coding", globalLibName + ".onLocationColouring")
      .addSeparator()
      .addItem("Perform some magic...", globalLibName + ".onPerformMagic")
      .addToUi();
  trace("< Custom rota menu added");
}

function addQuoteMenu() {
  trace("> Adding custom quote menu");
  let ui = SpreadsheetApp.getUi();
  ui.createMenu("Quote" + globalLibMenuTag)
      .addItem("Update Quote", globalLibName + ".onUpdateQuote")
      .addSeparator()
      .addItem("Create Estimate Summary", globalLibName + ".onCreateEstimateSummary")
      .addSeparator()
      .addItem("Perform some magic...", globalLibName + ".onPerformMagic")
      .addToUi();
  trace("< Custom quote menu added");
}

function addEnqiriesMenu() {
  trace("> Adding custom enquiries menu");
  let ui = SpreadsheetApp.getUi();
  ui.createMenu("Enquiries" + globalLibMenuTag)
      .addItem("Update Enquiries", globalLibName + ".onUpdateEnquiries")
      //.addItem("Create Quote/Client Sheet", globalLibName + ".onCreateNewClientSheet")
      .addItem("Open Quote/Client Sheet", globalLibName + ".onOpenClientSheet")
      .addItem("Prepare Client Document Structure", globalLibName + ".onPrepareClientStructure")
      .addSeparator()
      .addSubMenu(ui.createMenu("Test")
        .addItem("Test Case 1", globalLibName + ".onTestCase1")
        .addItem("Test Case 2", globalLibName + ".onTestCase2")
        .addItem("Test Case 3", globalLibName + ".onTestCase3")
      )
      .addToUi();
  trace("< Custom enquiries menu added");
}

function addWeddingPackagesMenu() {
  trace("> Adding custom wedding packages menu");
  let ui = SpreadsheetApp.getUi();
  ui.createMenu("Prices & Packages" + globalLibMenuTag)
      .addItem("Refresh Price List", globalLibName + ".onRefreshPriceList")
      .addItem("Update Packages", globalLibName + ".onUpdatePackages")
      .addToUi();
  trace("< Custom wedding packages menu added");
}

function addEmailMenu() {
  trace("> Adding custom email menu");
  let ui = SpreadsheetApp.getUi();
  let menu = ui.createMenu("Email" + globalLibMenuTag)
  .addItem("Draft selected email", globalLibName + ".onDraftSelectedEmail")
  .addSeparator();
  EmailTemplateList.singleton.populateMenu(menu)
  .addToUi();
  trace("< Custom email menu added");
}

addProductsSheetMenu
function addProductsSheetMenu() {
  trace("> Adding custom products menu");
  let ui = SpreadsheetApp.getUi();
  ui.createMenu("Products" + globalLibMenuTag)
      .addItem("Refresh Price List", globalLibName + ".onRefreshPriceList")
      .addItem("Update Packages", globalLibName + ".onUpdatePackages")
      .addSubMenu(ui.createMenu("Export")
                  .addItem("Clear Export Area", globalLibName + ".onPriceListClearExport")
                  .addItem("Clear Selection Ticks", globalLibName + ".onPriceListClearSelectionTicks")
                  .addItem("Export Ticked Items", globalLibName + ".onPriceListExportTicked")
                  .addItem("Export Marked Items", globalLibName + ".onPriceListExportSelection")
                  )
      .addSubMenu(ui.createMenu("Helpers")
                  .addItem("Test Case 1", globalLibName + ".onTestCase1")
                  .addItem("Test Case 2", globalLibName + ".onTestCase2")
                  )
      .addToUi();
  trace("< Custom products menu added");
}

