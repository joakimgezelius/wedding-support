//=============================================================================================
// Class Dialog

class Dialog {
  
  constructor() {
  }

  static prompt(title, message) {
    let ui = SpreadsheetApp.getUi();
    trace("Dialog.prompt " + title + ", " + message);
    let result = ui.prompt(title, message, ui.ButtonSet.OK_CANCEL);
    let button = result.getSelectedButton();
    let text = result.getResponseText();
    if (button == ui.Button.CANCEL) {
      trace("Dialog.prompt --> CANCEL");
      text = "CANCEL";
    } else {
      trace("Dialog.prompt --> " + text);
    }
    return text;
  }

  static confirm(title, message) {
    let ui = SpreadsheetApp.getUi();
    trace("Dialog.confirm " + title + ", " + message);
    let response = ui.alert(title, message, ui.ButtonSet.OK_CANCEL);
    let result = (response == ui.Button.OK) ? true : false;
    trace("Dialog.confirm --> " + result);
    return result;
  }

  static notify(title, message) {
    // https://developers.google.com/apps-script/reference/base/ui
    let ui = SpreadsheetApp.getUi();
    trace("Dialog.notify " + title + ", " + message);
    ui.alert(title, message, ui.ButtonSet.OK);
    trace("Dialog.notify");
  }

  static toast(message, title = null) {
    if (title == null) {
      SpreadsheetApp.getActive().toast(message);
    } else {
      SpreadsheetApp.getActive().toast(message, title);
    }
  }
}
