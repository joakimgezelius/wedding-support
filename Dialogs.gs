//=============================================================================================
// Class Dialog

var Dialog = function() {
}

Dialog.prompt = function(title, message) {
  var ui = SpreadsheetApp.getUi();
  trace("Dialog.prompt " + title + ", " + message);
  var result = ui.prompt(title, message, ui.ButtonSet.OK_CANCEL);
  var button = result.getSelectedButton();
  var text = result.getResponseText();
  if (button == ui.Button.CANCEL) {
    trace("Dialog.prompt --> CANCEL");
    text = "CANCEL";
  } else {
    trace("Dialog.prompt --> " + text);
  }
  return text;
}

Dialog.confirm = function(title, message) {
  var ui = SpreadsheetApp.getUi();
  trace("Dialog.confirm " + title + ", " + message);
  var response = ui.alert(title, message, ui.ButtonSet.OK_CANCEL);
  result = (response == ui.Button.OK) ? true : false;
  trace("Dialog.confirm --> " + result);
  return result;
}

Dialog.notify = function(title, message) {
  var ui = SpreadsheetApp.getUi();
  trace("Dialog.notify " + title + ", " + message);
  ui.alert(title, message, ui.ButtonSet.OK);
  trace("Dialog.notify");
}
