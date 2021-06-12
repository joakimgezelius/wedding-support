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
  //let sampleFolderId = "0B0i74Z2VzsKXNjdSUzIyZ0dSRjg";      // W & E's
  let sampleFolderId = "1SL8IXZ3bvU62yHXQTycUMKmD29F5tylT"; // 2020
  let folder = Folder.getById(sampleFolderId);
  let file1 = folder.getFile("Wedding Docs.pdf");
  let file2 = folder.getFile("foo");
  let folder1 = folder.getSubfolder("Welcome gift bag for couples");
  let folder2 = folder.getSubfolder("foo");
  folder.recursiveWalk();
}

function onTestCase2() {
  trace("onTestCase2");
  let user = User.active;
  trace("try again");
  user = User.active;
  //let clientTemplateFolderId = "1lIUlRJFAxoVsOy_Tdmga9ZqzTWZGDDxr";
  //let clientTemplateFolder = Folder.getById(clientTemplateFolderId);
  //clientTemplateFolder.listFiles();
  //clientTemplateFolder.recursiveWalk();
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
