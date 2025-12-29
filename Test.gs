function onTestCase1() {        // HubSpot Testing
  trace("onTestCase1");
  Spreadsheet.active.iterateOverNamedRanges();
  Spreadsheet.active.getRangeByName("AsanaAPIColumnIds");
  // HubSpotDataDictionary.importHubSpotTasks();
  // HubSpot.listContacts(); 
  // HubSpot.listDeals();
  // HubSpot.contactToDeal();
  // HubSpot.dealToContact();
  // HubSpot.listEngagement();
  // HubSpot.masterHubspot();
  // HubSpot.getClientData();
  // let spreadsheet = Spreadsheet.active;
  // let sheet1 = spreadsheet.getSheetByName("Crazy in love");
  // let range1 = sheet1.fullRange;
  // range1.setName("'Crazy in love'!TestRangeName");
  // let sheet2 = spreadsheet.getSheetByName("Banoffee Bonanza");
  // let range2= sheet2.fullRange;
  // range2.setName("'Banoffee Bonanza'!TestRangeName");
}

function onTestCase2() {    
  trace("onTestCase2");
  //get the right spreadsheet & sheet
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Deals');
  sheet.clear();
  sheet.appendRow(["Dealname", "Dealstage", "Close date", "Amount"]);

  //set up API letiables
  let offset = 0;
  let queryParams = '&limit=100&offset=' + offset + '&properties=dealname&properties=dealstage&properties=closedate&properties=amount'
  let url = 'https://api.hubapi.com/deals/v1/deal/paged?hapikey=d69a4027-d1cf-4730-9492-cf16faf333b1';
  let options = {
    "method": "GET",
    "muteHttpExceptions": true
  };

  //set up functions for call
  let apiCall = function(url, endpoint){
    let apiResponse = UrlFetchApp.fetch(url + endpoint,options);
    let json = JSON.parse(apiResponse);
    return json
  };

  let newDeals = apiCall(url, queryParams);
  trace(`Deals: ${newDeals}`);
}

function onTestCase3() {        // Other Testings
  trace("onTestCase3");
  //let hubSpotDataDictionary = HubSpotDataDictionary.current;
  //let hubSpotDataDictionary2 = HubSpotDataDictionary.current;
  let templateFolderLink = Spreadsheet.getCellValueLinkUrl("TemplateClientFolder");     // W & E's >> Upcoming
  let sourceFolder = Folder.getByUrl(templateFolderLink);
  let templateClientSheetLink = Spreadsheet.getCellValueLinkUrl("TemplateClientSheet"); // URL to the W & E's template sheet
  let templateClientSheet = Folder.getByUrl(templateClientSheetLink);
  let destinationFolderURL = Spreadsheet.getCellValueLinkUrl("ClientFoldersRoot"); // Destination - W & E's param named range ClientFoldersRoot
  let destinationFolder = Folder.getByUrl(destinationFolderURL);
  trace(`Source Folder: ${sourceFolder.trace}`);
  trace(`Destination Folder: ${destinationFolder.trace}`);
  destinationFolder.createShortcut(templateClientSheet);
  destinationFolder.createShortcut(sourceFolder, "New name for this shortcut");
}
