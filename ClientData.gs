class ClientData {

  static init() {
    if (ClientData.values === null) {
      trace("ClientData.init");
      ClientData.values = {};
      let clientData = Range.getByName("ClientData");
      let values = clientData.values;
      // Iterate over the named Client Data properties, and add an associative record for each non-empty value.
      values.forEach(item => 
        { 
          // Assumes id's are in colun 0, values in column 3
          let propertyId = item[0];
          let propertyValue = item[3];
          if (propertyId !== "" && propertyValue !== "") {
            //trace(`item: ${propertyId} = ${propertyValue}`);
            ClientData.values[propertyId] = propertyValue;
          }
        }
      );
      // Derived values go here:

    }
  }

  static lookup(propertyId) {
    ClientData.init();
    let propertyValue = ClientData.values[propertyId];
    trace(`ClientData.lookup[${propertyId}] --> ${propertyValue}`);
    return propertyValue;
  }


} // ClientData

ClientData.values = null; // Static property

function CLIENTDATA(propertyId) {
  return ClientData.lookup(propertyId);
}
