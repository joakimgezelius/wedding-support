
//=============================================================================================
// Class Rota
//

class Rota {

  onActivityColouring() {
    trace("Rota.onActivityColouring ");
  }

  onSupplierColouring() {
    trace("Rota.onSupplierColouring ");
  }

  onLocationColouring () {
    trace("Rota.onLocationColouring ");
  }

  onPerformMagic() {
    trace("Rota.onPerformMagic ");
  }

  onUpdateRota() {
    trace("Rota.onUpdateRota ");
    if (Dialog.confirm("Update Rota - Confirmation Required", "Are you sure you want to update the rota? It will overwrite the row numbers, make sure the sheet is sorted properly!") == true) {
      trace("Rota.onUpdateRota ");
//    let eventDetailsIterator = new EventDetailsIterator();
//    let eventDetailsUpdater = new EventDetailsUpdater();
//    eventDetailsIterator.iterate(eventDetailsUpdater);
    }
  }

  get trace() {
    return "{Rota}";
  }

}
