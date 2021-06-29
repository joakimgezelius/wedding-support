//----------------------------------------------------------------------------------------
// Wrapper for https://developers.hubspot.com/docs/api/crm/understanding-the-crm
//   Contacts   : https://developers.hubspot.com/docs/api/crm/contacts
//   Deals      : https://developers.hubspot.com/docs/api/crm/deals

class HubSpot {

  static listContacts() {
    let url = HubSpot.getUrl("contacts");
    let response = UrlFetchApp.fetch(url);
    trace(`HubSpot.listContacts --> ${response.getContentText()}`);
    let data = JSON.parse(response.getContentText());
    let results = data['results'];
    let paging = data.paging.next;
    let sheet = SpreadsheetApp.getActiveSheet();
    let header = ["Id", "Create Date", "Email", "First Name", "Object Id", "Last Modified Date", "Last Name", "Phone", "Website","Paging After","Paging Link"];
    let items = [header];
    results.forEach(function (result) {
    items.push([ result['properties'].hs_object_id, result['properties'].createdate, result['properties'].email, result['properties'].firstname, result['properties'].hs_object_id, result['properties'].lastmodifieddate, result['properties'].lastname, result['properties'].phone, result['properties'].website, paging.after, paging.link]);
    });
    sheet.getRange(3,1,items.length,items[0].length).setValues(items);
    results.forEach(function (result) {
      Logger.log(result['properties']);
     });
  }

  static listDeals() {
    let url = HubSpot.getUrl("deals");
    let response = UrlFetchApp.fetch(url);
    trace(`HubSpot.listContacts --> ${response.getContentText()}`);
    let data = JSON.parse(response.getContentText());
    let results = data['results'];
    let paging = data.paging.next;
    let sheet = SpreadsheetApp.getActiveSheet();
    let header = ["Amount", "Close Date", "Create Date", "Deal Name", "Deal Stage", "Last Modified Date", "HS Owner ID","Paging After","Paging Link"];
    let items = [header];
    results.forEach(function (result) {
    items.push([ result['properties'].amount, result['properties'].closedate, result['properties'].createdate, result['properties'].dealname, result['properties'].dealstage, result['properties'].hs_lastmodifieddate, result['properties'].hubspot_owner_id, paging.after, paging.link]);
    });
    sheet.getRange(3,1,items.length,items[0].length).setValues(items);
    results.forEach(function (result) {
      Logger.log(result['properties']);
     });
  }

  static contactToDeal() {
    let url = HubSpot.getUrl("contacts");
    let response = UrlFetchApp.fetch(url);
    trace(`HubSpot.contactToDeal --> ${response.getContentText()}`);
    let data = JSON.parse(response.getContentText());
    let results = data['results'];
    /*let sheet = SpreadsheetApp.getActiveSheet();
    let header = ["Id", "Create Date", "Email", "First Name", "Object Id", "Last Modified Date", "Last Name", "Phone", "Website"];
    let items = [header];
    results.forEach(function (result) {
    items.push([ result['properties'].hs_object_id, result['properties'].createdate, result['properties'].email, result['properties'].firstname, result['properties'].hs_object_id, result['properties'].lastmodifieddate, result['properties'].lastname, result['properties'].phone, result['properties'].website]);
    });
    sheet.getRange(3,1,items.length,items[0].length).setValues(items);*/
    results.forEach(function (result) {
      Logger.log(result['associations']);
    });
  }


  static getUrl(method) {
    //return `${HubSpot.baseUrl}/${method}?limit=100&hapikey=${HubSpot.key}`;       // for Contacts & Deals
    return `${HubSpot.baseUrl}/${method}?associations=deal&hapikey=${HubSpot.key}`; // for association of contact_to_deal
  }

} // User

HubSpot.baseUrl = "https://api.hubapi.com/crm/v3/objects";
HubSpot.key = "0020bf99-6b2a-4887-90af-adac067aacba";

