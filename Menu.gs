class Menu {

  static addEventMenu(libName) {
    trace("> Adding custom event menu");
    let ui = SpreadsheetApp.getUi();
    let menu = ui.createMenu("Event" + globalLibMenuTag)
      //.addItem("Clear Sheet", libName + ".onClearSheet")
      //.addSeparator()
      //.addItem("Pull Client Information", libName + ".onPullClientInformation") // In client.gs
      //.addSeparator()
      .addItem("Update Coordinator (no over-writes)", libName + ".onUpdateCoordinator")   // In Coordinator.gs
      .addItem("Force-Update Coordinator - over-writes data", libName + ".onUpdateCoordinatorForced")   // In Coordinator.gs
      .addItem("Check Client Sheet", libName + ".onCheckClientSheet")               // In ClientSheet.gs
      .addItem("Update Budget", libName + ".onUpdateBudget")    // In Budget.gs
      .addItem("Update Client Itinerary", libName + ".onUpdateClientItinerary")     // In itinerary.gs
      .addItem("Update Decor Summary", libName + ".onUpdateDecorSummary")           // In DecorSummary.gs
      .addItem("Update Supplier Costing", libName + ".onUpdateSupplierCosting")     // In SupplierCosting.gs      
      //.addItem("Update Supplier Venue Itinerary", libName + ".onUpdateSupplierVenueItinerary") // In VenueItinerary.gs
      //.addItem("Update Venue Itinerary", libName + ".onUpdateVenueItinerary")       // In VenueItinerary.gs
      .addItem("Pull Client Data From HubSpot", libName + ".onUpdateClientData")    // In HubSpotAPI.gs
      .addSeparator()
      //.addItem("Update Exchange Rates", libName + ".onUpdateExchangeRates")
      //.addSeparator()
      /*.addSubMenu(ui.createMenu("Prices")
                  .addItem("Import price list", libName + ".onImportPriceList")
                  .addItem("tbd...", libName + ".onImportPriceList")
                  )
      .addSeparator()*/
      //.addItem("Create New Client Sheet", libName + ".onCreateNewClientSheet")
      .addItem("Apply Format Template", libName + ".onApplyFormat")     // In SheetFormat.gs
      .addItem("Format Coordinator", libName + ".onFormatCoordinator"); // In SheetFormat.gs
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

  static addAsanaMenu(libName) {            // In AsanaAPI.gs
    trace("> Adding custom asana menu");
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
    trace("< Custom asana menu added");
  }

  static addTestMenu(libName) {
    trace("> Adding custom test menu for developer");
    let ui = SpreadsheetApp.getUi();
    let menu = ui.createMenu("Test" + globalLibMenuTag);
    Menu.addTestItems(libName, menu).addToUi();
    trace("< Custom  test menu added");
  }

  static addEventCoordinationMenu(libName) {
    trace("> Adding custom event coordination menu");
    let ui = SpreadsheetApp.getUi();
    let menu = ui.createMenu("Events" + globalLibMenuTag)
        .addItem("Update Event Coordination Sheet", libName + ".updateMasterQuery");
    Menu.addTestItems(libName, menu).addToUi();
    trace("< Custom event coordination menu added");
  }

  static addRotaMenu(libName) {           // In Rota.gs
    trace("> Adding custom rota menu");
    let ui = SpreadsheetApp.getUi();
    let menu = ui.createMenu("Rota" + globalLibMenuTag)
        .addItem("Update Rota Sheet", libName + ".onUpdateRotaSheet");
    Menu.addTestItems(libName, menu).addToUi();
    trace("< Custom rota menu added");
  }

  static addDecorPriceListMenu(libName) {         // In DecorPriceList.gs
    trace("> Adding custom decor price list menu");
    let ui = SpreadsheetApp.getUi();
    let menu = ui.createMenu("Decor" + globalLibMenuTag)
        .addItem("Update Decor Price List", libName + ".onUpdateDecorPriceList")
        .addItem("Update Decor Price List (NEW)", libName + ".onUpdateNewDecorPriceList");
    Menu.addTestItems(libName, menu).addToUi();
    trace("< Custom decor price list menu added");
  }

  static addShopSalesListMenu(libName) {         // In ShopSales.gs
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

  static addProjectsMenu(libName) {               // In Enquiries.gs
    trace("> Adding custom projects menu");
    let ui = SpreadsheetApp.getUi();
    let menu = ui.createMenu("Projects" + globalLibMenuTag)
        .addItem("Show Project Sidebar", libName + ".onShowProjectSidebar")
        .addItem("Update Enquiries", libName + ".onUpdateEnquiries")
        .addItem("Open Project Sheet", libName + ".onOpenProjectSheet")
        .addItem("Prepare New Project Document Structure", libName + ".onPrepareProjectStructure")
        .addSubMenu(ui.createMenu("Project Folder Maintenance")
          .addItem("Add payments folder & link", libName + ".onPreparePaymentsFolder")
          .addItem("Delete Project Document Structure", libName + ".onDeleteProjectDocumentStructure")
        );
        /*.addSubMenu(ui.createMenu("Prepare New Client Document Structure")
          .addItem("For Small Wedding/Event", libName + ".onPrepareClientStructureSmallWedding")
          .addItem("For Large Wedding/Event", libName + ".onPrepareClientStructureLargeWedding")        
        );*/
    Menu.addTestItems(libName, menu).addToUi();
    trace("< Custom projects menu added");
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
    trace("> Adding custom price list menu");
    let ui = SpreadsheetApp.getUi();
    let menu = ui.createMenu("Price List" + globalLibMenuTag)
        .addItem("Force-Update Price List", libName + ".onUpdatePriceListForced"); // PackagePriceList.gs
    Menu.addTestItems(libName, menu).addToUi();
    trace("< Custom price list menu added");
  }

  static addUtilitiesSheetMenu(libName) {         // In Utilities.gs
    trace("> Adding custom utilities menu");
    let ui = SpreadsheetApp.getUi();
    let menu = ui.createMenu("Utilities" + globalLibMenuTag)
        .addItem("Get File Info", libName + ".onGetFileInfo")                             // Utilities.gs
        .addItem("Get Folder Info", libName + ".onGetFolderInfo")                         // Utilities.gs
        .addItem("Move to Shared Drive", libName + ".onMoveToSharedDrive")                // Utilities.gs
        .addItem("Transfer File/Folder Ownership", libName + ".onTransferFileOwnership")  // Utilities.gs
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

