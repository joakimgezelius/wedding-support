//----------------------------------------------------------------------------------------
// Wrapper for https://developers.hubspot.com/docs/api/crm/understanding-the-crm
//   Contacts: https://developers.hubspot.com/docs/api/crm/contacts
//   Deals:    https://developers.hubspot.com/docs/api/crm/deals
//
class HubSpot {

  static listContacts() {
    let url = HubSpot.getUrl("contacts");
    let response = UrlFetchApp.fetch(url);
    trace(`HubSpot.listContacts --> ${response.getContentText()}`);
    let data = JSON.parse(response.getContentText());
    let results = data['results'];
  }

  static getUrl(method) {
    return `${HubSpot.baseUrl}/${method}?hapikey=${HubSpot.key}`;
  }

} // User

HubSpot.baseUrl = "https://api.hubapi.com/crm/v3/objects";
HubSpot.key = "0020bf99-6b2a-4887-90af-adac067aacba";


// Sample code from https://www.actiondesk.io/blog/google-script-query-hubspot-api-dashboard
//
function callHsapi() {
  var API_KEY = "YOUR_HUBSPOT_API_KEY"; // Replace YOUR_HUBSPOT_API_KEY with your API Key
  var url = "https://api.hubapi.com/crm/v3/objects/contacts?limit=10&archived=false&hapikey=YOUR_HUBSPOT_API_KEY"; // Replace YOUR_HUBSPOT_API_KEY with your API Key
  var response = UrlFetchApp.fetch(url);
  var data = JSON.parse(response.getContentText());
  var results = data['results'];
  var sheet = SpreadsheetApp.getActiveSheet();
  var header = ["Company", "Create Date", "Email", "First Name", "Last Modified Date", "Last Name", "Phone", "Website"];
  var items = [header];
  results.forEach(function (result) {
    items.push([result['properties'].company, result['properties'].createdate, result['properties'].email, result['properties'].firstname, result['properties'].lastmodifieddate, result['properties'].lastname, result['properties'].phone, result['properties'].website]);
  });
  sheet.getRange(1,1,items.length,items[0].length).setValues(items);
  results.forEach(function (result) {
    Logger.log(result['properties']);
  });
}
