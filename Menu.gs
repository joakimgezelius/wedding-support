class Menu {

  static addEventMenu(libName) {
    trace("> Adding custom event menu");
    let ui = SpreadsheetApp.getUi();
    let menu = ui.createMenu("Event" + globalLibMenuTag)
      //.addItem("Clear Sheet", libName + ".onClearSheet")
      //.addSeparator()
      //.addItem("Pull Client Information", libName + ".onPullClientInformation") // in client.gs
      //.addSeparator()
      .addItem("Update Coordinator (no over-writes)", libName + ".onUpdateCoordinator")
      .addItem("Force-Update Coordinator - over-writes data", libName + ".onUpdateCoordinatorForced")
      //.addItem("Check Coordinator", libName + ".onCheckCoordinator")
      .addItem("Update Budget", libName + ".onUpdateBudget")
      .addItem("Update Client Itinerary", libName + ".onUpdateClientItinerary") // In itinerary.gs
      .addItem("Update Decor Summary", libName + ".onUpdateDecorSummary")
      .addItem("Update Account Summary", libName + ".onUpdateAccountSummary")
      .addItem("Update Supplier Costing", libName + ".onUpdateSupplierCosting")
      //.addItem("Update Supplier Itinerary", libName + ".onUpdateSupplierItinerary")
      .addItem("Update Client Data", libName + ".onUpdateClientData")
      //.addItem("Update Staff Itinerary", libName + ".onUpdateStaffItinerary")
      //.addItem("Update Rota", libName + ".onUpdateRota")
      .addSeparator()
      .addItem("Update Exchange Rates", libName + ".onUpdateExchangeRates")
      .addSeparator()
      .addSubMenu(ui.createMenu("Prices")
                  .addItem("Import price list", libName + ".onImportPriceList")
                  .addItem("tbd...", libName + ".onImportPriceList")
                  )
      .addSeparator()
      //.addItem("Create New Client Sheet", libName + ".onCreateNewClientSheet")
      .addItem("Apply Format Template", libName + ".onApplyFormat")    // In SheetFormat.gs
      .addItem("Format Coordinator", libName + ".onFormatCoordinator");
//      .addSeparator()
//      .addSubMenu(ui.createMenu("Email")
//                  .addItem("Prepare client email draft 1", libName + ".onCreateFirstEmailDraft")
//                  .addItem("Send client email 1", libName + ".onSendFirstEmail")
//                  )
//    .addSubMenu(ui.createMenu("Tasks")
//                  .addItem("Small Wedding", libName + ".onShowSmallWeddingTasks")
//                  .addItem("Big Wedding", libName + ".onShowBigWeddingTasks")
//                  .addItem("Show All", libName + ".onShowAllTasks")
//                  )
  Menu.addTestItems(libName, menu).addToUi();
  trace("< Custom event menu added");
}

static addAsanaMenu(libName) {
    trace("> Adding custom event menu");
    let ui = SpreadsheetApp.getUi();
    let menu = ui.createMenu("Asana" + globalLibMenuTag)
      .addSubMenu(ui.createMenu("Project")
                  .addItem("Make the Project", libName + ".onCreateProject")
                  .addItem("Update the Project", libName + ".onUpdateProject")
                  .addItem("Delete the Project", libName + ".onDestroyProject")
                  )
      .addSubMenu(ui.createMenu("Task")
                  .addItem("Upload Tasks", libName + ".onCreateTask")
                  .addItem("Update Tasks", libName + ".onUpdateTask")
                  .addItem("Delete Tasks", libName + ".onDestroyTask")
                  )
  Menu.addTestItems(libName, menu).addToUi();
  trace("< Custom event menu added");
}

static addTestMenu(libName) {
  trace("> Adding custom test menu if user is developer");
  if (User.active.isDeveloper) {
    let ui = SpreadsheetApp.getUi();
    ui.createMenu("Test" + globalLibMenuTag)
/*
    .addSubMenu(ui.createMenu("Testing")
      .addItem("Clear", libName + ".onClearCoordinatorSummary")
      .addItem("List All Sheets", libName + ".onListAllSheets")
      .addItem("List Coordinator Sheets", libName + ".onListCoordinatorSheets")
      .addItem("Test Coordinator Summary", libName + ".onUpdateCoordinatorSummary")
      .addItem("Show Details", libName + ".onShowDetails")
      .addItem("Hide Details", libName + ".onHideDetails")
      .addSeparator()
    )
    */
      .addToUi();
    trace("< Custom  test menu added");
  }
}

static addEventCoordinationMenu(libName) {
  trace("> Adding custom event coordination menu");
  let ui = SpreadsheetApp.getUi();
  let menu = ui.createMenu("Events" + globalLibMenuTag)
      .addItem("Update Coordination Sheet", libName + ".onUpdateCoordinationSheet");
  Menu.addTestItems(libName, menu).addToUi();
  trace("< Custom event coordination menu added");
}

static addRotaMenu(libName) {
  trace("> Adding custom rota menu");
  let ui = SpreadsheetApp.getUi();
  let menu = ui.createMenu("Rota" + globalLibMenuTag)
      .addItem("Update Rota Sheet", libName + ".onUpdateRotaSheet");
  Menu.addTestItems(libName, menu).addToUi();
  trace("< Custom rota menu added");
}

static addDecorPriceListMenu(libName) {
  trace("> Adding custom decor price list menu");
  let ui = SpreadsheetApp.getUi();
  let menu = ui.createMenu("Decor" + globalLibMenuTag)
      .addItem("Update Decor Price List", libName + ".onUpdateDecorPriceList")
      .addItem("Update Decor Price List (NEW)", libName + ".onUpdateNewDecorPriceList");
  Menu.addTestItems(libName, menu).addToUi();
  trace("< Custom decor price list menu added");
}

static addShopSalesListMenu(libName) {
  trace("> Adding custom shop sales & stock list menu");
  let ui = SpreadsheetApp.getUi();
  let menu = ui.createMenu("Shop (Sales & Stock)" + globalLibMenuTag)
      .addItem("Update Daily Sales", libName + ".onUpdateDailySales")
      .addItem("Update Shop Stock", libName + ".onUpdateShopStock");
  Menu.addTestItems(libName, menu).addToUi();
  trace("< Custom shop sales & stock list menu added");
}

static addQuoteMenu(libName) {
  trace("> Adding custom quote menu");
  let ui = SpreadsheetApp.getUi();
  let menu = ui.createMenu("Quote" + globalLibMenuTag)
      .addItem("Update Quote", libName + ".onUpdateQuote")
      .addSeparator()
      .addItem("Create Estimate Summary", libName + ".onCreateEstimateSummary");
  Menu.addTestItems(libName, menu).addToUi();
  trace("< Custom quote menu added");
}

static addEnqiriesMenu(libName) {
  trace("> Adding custom enquiries menu");
  let ui = SpreadsheetApp.getUi();
  let menu = ui.createMenu("Enquiries" + globalLibMenuTag)
      .addItem("Update Enquiries", libName + ".onUpdateEnquiries")
      //.addItem("Create Quote/Client Sheet", libName + ".onCreateNewClientSheet")
      .addItem("Open Quote/Client Sheet", libName + ".onOpenClientSheet")
      .addSubMenu(ui.createMenu("Prepare New Client Document Structure")
        .addItem("For Small Wedding/Event", libName + ".onPrepareClientStructureSmallWedding")
        .addItem("For Large Wedding/Event", libName + ".onPrepareClientStructureLargeWedding")
      );
  Menu.addTestItems(libName, menu).addToUi();
  trace("< Custom enquiries menu added");
}

  static addWeddingPackagesMenu(libName) {
    trace("> Adding custom wedding packages menu");
    let ui = SpreadsheetApp.getUi();
    let menu = ui.createMenu("Prices & Packages" + globalLibMenuTag)
        .addItem("Refresh Price List", libName + ".onRefreshPriceList")
        .addItem("Update Packages", libName + ".onUpdatePackages");
    Menu.addTestItems(libName, menu).addToUi();
    trace("< Custom wedding packages menu added");
  }

  static addEmailMenu(libName) {
    trace("> Adding custom email menu");
    let ui = SpreadsheetApp.getUi();
    let menu = ui.createMenu("Email" + globalLibMenuTag)
    .addItem("Draft selected email", libName + ".onDraftSelectedEmail")
    .addSeparator();
    EmailTemplateList.singleton.populateMenu(menu);
    Menu.addTestItems(libName, menu).addToUi();
    trace("< Custom email menu added");
  }

  static addProductsSheetMenu(libName) {
    trace("> Adding custom products menu");
    let ui = SpreadsheetApp.getUi();
    let menu = ui.createMenu("Products" + globalLibMenuTag)
        .addItem("Force-Update Price List", libName + ".onUpdatePriceListForced")
        .addItem("Refresh Price List", libName + ".onRefreshPriceList")
        .addItem("Update Packages", libName + ".onUpdatePackages")
        .addSubMenu(ui.createMenu("Export")
                    .addItem("Clear Export Area", libName + ".onPriceListClearExport")
                    .addItem("Clear Selection Ticks", libName + ".onPriceListClearSelectionTicks")
                    .addItem("Export Ticked Items", libName + ".onPriceListExportTicked")
                    .addItem("Export Marked Items", libName + ".onPriceListExportSelection")
                    );
    Menu.addTestItems(libName, menu).addToUi();
    trace("< Custom products menu added");
  }

  static addUtilitiesSheetMenu(libName) {
    trace("> Adding custom utilities menu");
    let ui = SpreadsheetApp.getUi();
    let menu = ui.createMenu("Utilities" + globalLibMenuTag)
        .addItem("Get file/folder info", libName + ".onGetFileInfo")
        .addItem("Move to Shared Drive", libName + ".onMoveToSharedDrive")
    Menu.addTestItems(libName, menu).addToUi();
    trace("< Custom products menu added");
  }

  static addTestItems(libName, menu) {
    return /*User.active.isDeveloper*/ true ? 
      menu
        .addSeparator()
        .addSubMenu(SpreadsheetApp.getUi().createMenu("Test (for developers)")
        .addItem("Test Case 1", libName + ".onTestCase1")
        .addItem("Test Case 2", libName + ".onTestCase2")
        .addItem("Test Case 3", libName + ".onTestCase3")
      ) : menu;
  }

}

