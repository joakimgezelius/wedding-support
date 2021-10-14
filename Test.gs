function onTestCase1() {        // HubSpot Testing
  trace("onTestCase1");  
  //HubSpot.listContacts(); 
  HubSpot.listDeals();
  //HubSpot.contactToDeal();
  //HubSpot.listEngagement();
  //HubSpot.masterHubspot();
}

function onTestCase2() {        // Asana Testing
  trace("onTestCase2");
  //Task.update();
  //Task.destroy(); 
  //Subtask.create();
  //Asana.getTaskNames();
  //Asana.getTaskGid();
  //Asana.getProjectSections();
}

function onTestCase3() {        // Other Testings
  trace("onTestCase3");
  HubSpot.contactToDeal();
}

