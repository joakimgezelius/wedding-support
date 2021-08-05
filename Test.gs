function onTestCase1() {        // HubSpot Testing
  trace("onTestCase2");  
  //HubSpot.listContacts(); 
  //HubSpot.listDeals();
  //HubSpot.contactToDeal();
  //HubSpot.listEngagement();
  HubSpot.masterHubspot();
}

function onTestCase2() {        // Asana Testing
  trace("onTestCase2");
  //Project.create();
  Task.create();
}

function onTestCase3() {        // Other Testings
  trace("onTestCase3");
  let user = User.active;
  trace(`${user}`);
}
