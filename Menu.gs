// var ui = SpreadsheetApp.getUi();

function addWeddingMenu() {
  trace("> Adding custom wedding menu");
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Wedding")
      //.addItem("Create New Client Sheet", globalLibName + ".onCreateNewClient")
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
      .addItem("Update Supplier Itinerary", globalLibName + ".onUpdateSupplierItinerary")
      //.addItem("Update Rota", globalLibName + ".onUpdateRota")
      .addSeparator()
//      .addSubMenu(ui.createMenu("Tasks")
//                  .addItem("Small Wedding", globalLibName + ".onShowSmallWeddingTasks")
//                  .addItem("Big Wedding", globalLibName + ".onShowBigWeddingTasks")
//                  .addItem("Show All", globalLibName + ".onShowAllTasks")
//                  )
      .addSubMenu(ui.createMenu("Helpers")
                  .addItem("Set Colour", globalLibName + ".onSetColour")
                  .addItem("Test Case 1", globalLibName + ".onTestCase1")
                  .addItem("Test Case 2", globalLibName + ".onTestCase2")
                  )
      .addSeparator()
      .addItem("Update Exchange Rates", globalLibName + ".onUpdateExchangeRates")
      .addSeparator()
      .addSubMenu(ui.createMenu("Email")
                  .addItem("Prepare client email draft 1", globalLibName + ".onCreateFirstEmailDraft")
                  .addItem("Send client email 1", globalLibName + ".onSendFirstEmail")
                  )
      .addToUi();
  trace("< Custom wedding menu added");
}

function addTestMenu() {
  trace("> Adding custom test menu");
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Test")
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
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Rota")
      .addItem("Update Rota", "Rota.onUpdateRota")
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
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Quote")
      .addItem("Update Quote", globalLibName + ".onUpdateQuote")
      .addSeparator()
      .addItem("Create Estimate Summary", globalLibName + ".onCreateEstimateSummary")
      .addSeparator()
      .addItem("Perform some magic...", globalLibName + ".onPerformMagic")
      .addToUi();
  trace("< Custom quote menu added");
}
