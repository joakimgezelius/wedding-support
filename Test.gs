function onTestCase1() {        // HubSpot Testing
  trace("onTestCase1");  
  //HubSpot.listContacts(); 
  HubSpot.listDeals();
  //HubSpot.contactToDeal();
  //HubSpot.dealToContact();
  //HubSpot.listEngagement();
  //HubSpot.masterHubspot();
  //HubSpot.getClientData();
}

function onTestCase2() {    
  trace("onTestCase2");
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Coordinator");
  let values = sheet.getRange("A6:AJ").getValues();
  let n = 0;
  for (n; n < values.length; n++) {
    if(values[n][0]=="") { 
      n++;
      break
    }
  }
  trace(`Row : ${n+5}`);
}

function onTestCase3() {        // Other Testings
  trace("onTestCase3");
  let hubSpotDataDictionary = HubSpotDataDictionary.current;
  let hubSpotDataDictionary2 = HubSpotDataDictionary.current;
}

