/*function onTestCase1() {
  trace("onTestCase1");
  //Error.break;
  
  let spreadsheet = Spreadsheet.active;
  let sheet = spreadsheet.getSheetByName("Suppliers Price List");
  let selection = sheet.selection;
  let range = selection.getActiveRange();
  Dialog.notify("Selection", range.getA1Notation());
  //Error.break;
}*/

function onTestCase1() {
  trace("onTestCase1");
  let srcFolderId = "10XxkUmEccJ73aKaznSZITiqx2BvaBPJz";    //  Client Template/Office Use
  let dstFolderId = "19y3-Zou_RAWHZKaZ_5W_FJXql_Pz-gdd";    //  W & E's
  weddingEventsFolder.recursiveWalk();
}

function onTestCase2() {
  trace("onTestCase2");
  let clientTemplateFolderId = "1lIUlRJFAxoVsOy_Tdmga9ZqzTWZGDDxr";
  let clientTemplateFolder = Folder.getById(clientTemplateFolderId);
  //clientTemplateFolder.listFiles();
  clientTemplateFolder.recursiveWalk();
}

function onTestCase3() {
  trace("onTestCase3");
  let sourceFolderId = "1lIUlRJFAxoVsOy_Tdmga9ZqzTWZGDDxr";        // Client Template
  let destinationFolderId = "19y3-Zou_RAWHZKaZ_5W_FJXql_Pz-gdd";   // W & E's
  let sourceFolder = Folder.getById(sourceFolderId);
  let destinationFolder = Folder.getById(destinationFolderId);
  //sourceFolder.listFolders();
  sourceFolder.copyTo(destinationFolder, "New Name");
}




