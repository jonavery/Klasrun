function doGet() {
  return getXML();
}

function getXML() {
  var xsi = XmlService.getNamespace("xsi", "http://www.w3.org/2001/XMLSchema-instance");
  var noNS = XmlService.getNamespace("noNamespaceSchemaLocation", "amzn-envelope.xsd");
  
  var Header = XmlService.createElement('Header')
    .addContent(XmlService.createElement('DocumentVersion').setText("1.01"))
    .addContent(XmlService.createElement('MerchantIdentifier').setText('A3FA9W3CDIWR8F'));
  var Envelope = XmlService.createElement('AmazonEnvelope')
    .setAttribute("noNamespaceSchemaLocation", "amzn-envelope.xsd", xsi)
    .addContent(Header)
    .addContent(XmlService.createElement('MessageType').setText('CartonContentsRequest'));
  
  // Initialize product sheet and count products to be shipped.
  var sheetMWS = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('MWS');
  var products = sheetMWS.sort(12).getDataRange().getValues();
  var counts = {};
  for (var i = 1; i < products.length; i++) {
    counts[products[i][11]] = 1 + (counts[products[i][11]] || 0);
  }
  
  // f = index of first item in shipment
  // k = index of shipment in counts array
  // j = cartonid number
  // i = index of current item within products array
  var f = 1;
  for (var k = 0; k < counts.length; k++) {
    var ShipmentId = products[f][11];
  
    var CartonContentsRequest = XmlService.createElement('CartonContentsRequest')
      .addContent(XmlService.createElement('ShipmentId').setText(ShipmentId))
      .addContent(XmlService.createElement('NumCartons').setText(String(counts[k])));
  
    var j = 0;
    for (var i=f; i < f+counts[k]; i++) {
      var Item = XmlService.createElement('Item')
        .addContent(XmlService.createElement('SKU').setText(products[i][0]))
        .addContent(XmlService.createElement('QuantityShipped').setText('1'))
        .addContent(XmlService.createElement('QuantityInCase').setText('1'));
      j++;
      var Carton = XmlService.createElement('Carton')
        .addContent(XmlService.createElement('CartonId').setText(String(j)))
        .addContent(Item);
      CartonContentsRequest.addContent(Carton);
    }

    var Message = XmlService.createElement('Message')
      .addContent(XmlService.createElement('MessageID').setText(String(k+1)))
      .addContent(CartonContentsRequest);
    Envelope.addContent(Message);
    f += counts[k];
  }
  
  var document = XmlService.createDocument(Envelope);
  var xml = XmlService.getPrettyFormat().format(document);
  Logger.log(xml)
  return ContentService.createTextOutput(xml);
}
