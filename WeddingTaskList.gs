var WeddingType = { small:"small", big:"big" };


function onShowSmallWeddingTasks() {
  trace("onShowSmallWeddingTasks");
  var weddingTaskList = new WeddingTaskList();
  weddingTaskList.showTasks(WeddingType.small);
}

function onShowBigWeddingTasks() {
  trace("onShowBigWeddingTasks");
  var weddingTaskList = new WeddingTaskList();
  weddingTaskList.showTasks(WeddingType.big);
}

function onShowAllTasks() {
  trace("onShowAllTasks");
  var weddingTaskList = new WeddingTaskList();
  weddingTaskList.showAllTasks();
}


//=============================================================================================
// Class WeddingTaskList

var WeddingTaskList = function() {
  this.column = { always:1, small:2, big:3 };
  this.range = Range.getByName("WeddingTaskList");
  this.rowCount = this.range.getHeight();
  this.sheet = this.range.getSheet();
//  EventRow[columnName] = column;
  trace("NEW " + this.trace());
}


WeddingTaskList.prototype.showTasks = function(weddingType) {
  switch (weddingType) {
    case WeddingType.small: selectorColumn = this.column.small;  break;
    case WeddingType.big:   selectorColumn = this.column.big;    break;
  }
  trace("WeddingTaskList.showTasks type: " + weddingType + " selector column: " + selectorColumn);
  this.showAllTasks();
  this.sheet.hideColumns(1, 3);
  for (var rowOffset = 0; rowOffset < this.rowCount; rowOffset++) {
    var alwaysShowCell = this.range.getCell(1+rowOffset, this.column.always);
    var selectorCell = this.range.getCell(1+rowOffset, selectorColumn);
    var absoluteRowNumber = this.range.getRow()+rowOffset;
    if (!alwaysShowCell.getValue() && !selectorCell.getValue()) {
//    trace("Hide row " + rowOffset);    
      this.sheet.hideRows(absoluteRowNumber);
    }
  }
}


WeddingTaskList.prototype.showAllTasks = function() {
  trace("WeddingTaskList.showAllTasks");
  this.sheet.showRows(this.range.getRow(), this.rowCount);
  this.sheet.showColumns(1, 3);
}


WeddingTaskList.prototype.trace = function() {
  return "{WeddingTaskList range=" + Range.trace(this.range) + ", rowCount=" + this.rowCount + "}";
}

