/*function onTestCase1() {
  trace("onTestCase1");
  //Error.break;
  
  let spreadsheet = Spreadsheet.active;
  let sheet = spreadsheet.getSheetByName("Suppliers Price List");
  let selection = sheet.selection;
  let range = selection.getActiveRange();
  Dialog.notify("Selection", range.getA1Notation());
  //Error.break;
   trace("onTestCase1");
  let srcFolderId = "10XxkUmEccJ73aKaznSZITiqx2BvaBPJz";    //  Client Template/Office Use
  //let sampleFolderId = "0B0i74Z2VzsKXNjdSUzIyZ0dSRjg";      // W & E's
  let sampleFolderId = "1SL8IXZ3bvU62yHXQTycUMKmD29F5tylT"; // 2020
  let folder = Folder.getById(sampleFolderId);
  let file1 = folder.getFile("Wedding Docs.pdf");
  let file2 = folder.getFile("foo");
  let folder1 = folder.getSubfolder("Welcome gift bag for couples");
  let folder2 = folder.getSubfolder("foo");
  folder.recursiveWalk();
}*/

function onTestCase1() {
  trace("onTestCase2");
  // /HubSpot.listContacts(); 
  HubSpot.masterHubspot();
}

function onTestCase2() {
  trace("onTestCase2");
  HubSpot.listDeals();
}

function onTestCase3() {
  trace("onTestCase3");
  //HubSpot.contactToDeal();
  HubSpot.listEngagement();
  //let user = User.active;
  //trace(`${user}`);
}
