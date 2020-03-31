
function onPullClientInformation() {
  trace("onPullClientInformation");
  if (Dialog.confirm("Pull Client Information - Confirmation Required", "Are you sure you want to pull client information from form database? It will overwrite existing client information!") == true) {
    var clientIdRange = Range.getByName("ClientId");
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


var Client = function () {
  Client.currentClient = this;
  this.name = "UNDEFINED";
}


Client.databaseSheetId = "1E-dE-S1mAXSCGcnshG9mBF8h76RAF0QwABlyAIKjHbE";
Client.databaseSheet = null;
Client.databaseRange = null;
Client.currentClient = null;

Client.prototype.trace = function() {
  return "{Client " + this.name + "}";
}

Client.LookupById = function(clientId) {
  var client = new Client;
  client.id = clientId;
  client.name = "Some Name";
  return client;
}
