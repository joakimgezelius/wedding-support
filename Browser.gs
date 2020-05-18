class Browser {

  static newTab(url) {
    let js = " \
      <script> \
          window.open('" + url + "'); \
          google.script.host.close(); \
      </script> \
      ";
    let html = HtmlService.createHtmlOutput(js)
    .setHeight(10)
    .setWidth(100);
    SpreadsheetApp.getUi().showModalDialog(html, 'Loading...');
  }
  
}
