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
  let url = "https://api.hubapi.com/crm/v3/objects/deals?limit=100&properties=hs_object_id,amount,closedate,createdate,dealname,description,hubspot_owner_id,dealstage,dealtype,departure_date,hs_forecast_amount,hs_manual_forecast_category,hs_forecast_probability,hubspot_team_id,hs_lastmodifieddate,hs_next_step,num_associated_contacts,hs_priority,pipeline&archived=false&hapikey=0020bf99-6b2a-4887-90af-adac067aacba";
  let response = UrlFetchApp.fetch(url);
  let data = JSON.parse(response.getContentText());
  let results = data['results'];
  let paging = data.paging.next;
  let sheet = SpreadsheetApp.getActiveSheet();
  let header = ["Deal ID", "Amount", "Close Date", "Create Date", "Deal Name", "Deal Description", "Deal Owner", "Deal Type", "Deal Stage", "Departure Date", "Forecast Amount", "Forecast Category", "Forecast Probabilty", "HubSpot Team", "Last Modified Date", "Next Step", "Number of Contacts", "Priority", "Pipeline"];
  let items = [header];
  results.forEach(function (result) {
  if(result['properties'].dealstage !== "closedlost") {
    items.push([ result['properties'].hs_object_id, result['properties'].amount, result['properties'].closedate, result['properties'].createdate, result['properties'].dealname, result['properties'].description, result['properties'].hubspot_owner_id, result['properties'].dealtype, result['properties'].dealstage, result['properties'].departure_date, result['properties'].hs_forecast_amount, result['properties'].hs_manual_forecast_category, result['properties'].hs_forecast_probability, result['properties'].hubspot_team_id, result['properties'].hs_lastmodifieddate, result['properties'].hs_next_step, result['properties'].num_associated_contacts, result['properties'].priority, result['properties'].pipeline]);      
     }
  });
  sheet.getRange(3,1,items.length,items[0].length).setValues(items);
  let apiCall = function(url) {
    let response = UrlFetchApp.fetch(url);
    let data = JSON.parse(response.getContentText());
    return data;
  };
  do {
    let nextData = apiCall(paging.link+"&hapikey=0020bf99-6b2a-4887-90af-adac067aacba");
    Logger.log(nextData);
  } while (paging.after != null);
}

function getFirstEmptyRowWholeRow() {
  let sheet = SpreadsheetApp.getActiveSheet();
  let range = sheet.getDataRange();
  let values = range.getValues();
  let row = 0;
  for ( row; row < values.length; row++) {
    if (!values[row].join("")) break;
  }
  return (row+1);
}

function onTestCase3() {        // Other Testings
  trace("onTestCase3");
  let hubSpotDataDictionary = HubSpotDataDictionary.current;
  let hubSpotDataDictionary2 = HubSpotDataDictionary.current;
}

