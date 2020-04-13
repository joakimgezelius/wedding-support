function onUpdateItinerary() {
  trace("onUpdateItinerary");
//updatePagedItinerary();
  updateDynamicItinerary();
}

function updatePagedItinerary() {
  trace("updatePagedItinerary");
  var eventDetailsIterator = new EventDetailsIterator();
  var itineraryBuilder = new StaticItineraryBuilder("Itinerary");
  eventDetailsIterator.sort(SortType.time);
  eventDetailsIterator.iterate(itineraryBuilder);
}

function updateDynamicItinerary() {
  trace("updateDynamicItinerary");
  var eventDetailsIterator = new EventDetailsIterator();
  var itineraryBuilder = new DynamicItineraryBuilder("DynamicItinerary");
  eventDetailsIterator.sort(SortType.time);
  eventDetailsIterator.iterate(itineraryBuilder);
}


//=============================================================================================
// Class StaticItineraryBuilder
//
var StaticItineraryBuilder = function(targetRangeName) {
  this.targetRange = Range.getByName(targetRangeName);
  this.maxTargetRows = this.targetRange.getHeight();
  this.targetRowOffset = 0;
  trace("NEW " + this.trace());
}

StaticItineraryBuilder.prototype.onBegin = function() {
  trace("StaticItineraryBuilder.onBegin - reset context " + this.trace());
  this.targetRowOffset = 0;
  Range.clear(this.targetRange);
}
  
StaticItineraryBuilder.prototype.onEnd = function() {
  trace("StaticItineraryBuilder.onEnd - no-op");
}

StaticItineraryBuilder.prototype.onTitle = function(row) {
  trace("StaticItineraryBuilder.onTitle - ignore row: " + row.getTitle());
}

StaticItineraryBuilder.prototype.onRow = function(row) {
  if (row.isItineraryTicked()) { // This is an itinerary item
    trace("StaticItineraryBuilder.onRow Ticked: " + row.getDescription());
    var targetRow = this.getNextTargetRow();
    //trace("StaticItineraryBuilder.onRow got target row: " + Range.trace(targetRow));
    var column = 1;
    targetRow.getCell(1,column++).setValue(row.getDate());
    targetRow.getCell(1,column++).setValue(row.getTime());
    targetRow.getCell(1,column++).setValue(row.getLocation());
    targetRow.getCell(1,column++).setValue(row.getDescription());
  } else {
    trace("StaticItineraryBuilder.onRow Unticked, ignore: " + row.getDescription());
  }
}

// private method getNextTargetRow
//
StaticItineraryBuilder.prototype.getNextTargetRow = function() {
  return this.targetRange.offset(Math.min(this.targetRowOffset++, this.maxTargetRows-1), 0, 1); // A range of 1 row height
}
  
StaticItineraryBuilder.prototype.trace = function() {
  return "{StaticItineraryBuilder " + Range.trace(this.targetRange) + "}";
}


//=============================================================================================
// Class DynamicItineraryBuilder, subclassing StaticItineraryBuilder 
//
var DynamicItineraryBuilder = function(targetRangeName) {
  StaticItineraryBuilder.call(this, targetRangeName); // Call base class constructor
  trace("NEW " + this.trace());
}

// (Subclassing wiring........
DynamicItineraryBuilder.prototype = Object.create(StaticItineraryBuilder.prototype); // Add base class prototype  
Object.defineProperty(DynamicItineraryBuilder.prototype, "constructor", {
  value: DynamicItineraryBuilder, 
  enumerable: false, // so that it does not appear in "for in" loop
  writable: true });
//......... Subclassing wiring)
  
DynamicItineraryBuilder.prototype.onBegin = function() {
  trace("DynamicItineraryBuilder.onBegin - reset context " + this.trace());
  StaticItineraryBuilder.prototype.onBegin.call(this) // Call base class method
  // Delete all but the first and the last row in the target range
  var targetRangeHeight = this.targetRange.getHeight();
  if (targetRangeHeight > 2) {
    this.targetRange.getSheet().deleteRows(this.targetRange.getRowIndex() + 1, targetRangeHeight - 2);
  }
} 
  
// private method getNextTargetRow
//
DynamicItineraryBuilder.prototype.getNextTargetRow = function() {
  var targetRow = this.targetRange.offset(this.targetRowOffset++, 0, 1); // A range of 1 row height
  targetRow.getSheet().insertRowAfter(targetRow.getRowIndex());
  return targetRow;
}
  
DynamicItineraryBuilder.prototype.trace = function() {
  return "{DynamicItineraryBuilder " + Range.trace(this.targetRange) + "}";
}
