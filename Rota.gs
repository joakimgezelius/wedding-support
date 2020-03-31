
//=============================================================================================
// Class Rota
//

var Rota = function() {
}

Rota.onActivityColouring = function() {
  trace("Rota.onActivityColouring ");
}

Rota.onSupplierColouring = function() {
  trace("Rota.onSupplierColouring ");
}

Rota.onLocationColouring = function() {
  trace("Rota.onLocationColouring ");
}

Rota.onPerformMagic = function() {
  trace("Rota.onPerformMagic ");
}

Rota.onUpdateRota = function() {
  trace("Rota.onUpdateRota ");
  if (Dialog.confirm("Update Rota - Confirmation Required", "Are you sure you want to update the rota? It will overwrite the row numbers, make sure the sheet is sorted properly!") == true) {
    trace("Rota.onUpdateRota ");
    //var eventDetailsIterator = new EventDetailsIterator();
    //var eventDetailsUpdater = new EventDetailsUpdater();
    //eventDetailsIterator.iterate(eventDetailsUpdater);
  }
}

Rota.trace = function(rota) {
  return "{" + rota + "}";
}
