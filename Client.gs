
function onPullClientInformation() {
  trace("onPullClientInformation");
  if (Dialog.confirm("Pull Client Information - Confirmation Required", "Are you sure you want to pull client information from form database? It will overwrite existing client information!") == true) {
    var clientId = Spreadsheet.getCellValue("ClientId");
    var client = Client.LookupById("foobar");
    trace("Current client: " + client.trace());
    Client.databaseSheet = SpreadsheetApp.openById(Client.databaseSheetId);
    Client.databaseRange = Range.getByName("ClientDatabase", Client.databaseSheet);
    Client.databaseColumnTagsRange = Range.getByName("ClientDatabaseColumnTags", Client.databaseSheet);
    tags = Client.databaseColumnTagsRange;
    trace(Client.databaseSheet.getName() + " " + Range.trace(Client.databaseRange));
  }
}


//=============================================================================================
// Class Client


class Client {
  
  constructor() {
    Client.currentClient = this;
    this.name = "UNDEFINED";
  }



  static LookupById(clientId) {
    var client = new Client;
    client.id = clientId;
    client.name = "Some Name";
    return client;
  }

  get trace() {
    return "{Client " + this.name + "}";
  }

}

Client.databaseSheetId = "1E-dE-S1mAXSCGcnshG9mBF8h76RAF0QwABlyAIKjHbE";
Client.databaseSheet = null;
Client.databaseRange = null;
Client.currentClient = null;
