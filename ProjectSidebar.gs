

function onShowProjectSidebar() {
  var widget = HtmlService.createHtmlOutputFromFile("ProjectSidebarWidget.html");
  widget.setTitle("Project Details (test)");
  SpreadsheetApp.getUi().showSidebar(widget);
}

function displayToast() {
  SpreadsheetApp.getActive().toast("Hi there!");
}