//======================================================================================
// Experimental below

function onShowDetails() {
  showDetails(true);
}

function onHideDetails() {
  showDetails(false);
}


function showDetails(doShow) {
  trace("showDetails");
  var rows = targetRange.getHeight();
  for (var row = 0; row < rows; row++) {
    var targetRow = targetRange.offset(row, 0, 1);
    if (true) {
      if (doShow == true)
        targetSheet.showRow(targetRow);
      else
        targetSheet.hideRow(targetRow);    
    }   
  }
}
