
function onCheckClientSheet() {
  trace("onCheckClientSheet");

  // Check the coordinator
  let eventDetails = new EventDetails();
  let eventDetailsChecker = new EventDetailsChecker(); // In Coordinator.gs
  eventDetails.apply(eventDetailsChecker);
}
