function doGet() {
  return getXML();
}

function getXML() {
  var xsi = XmlService.getNamespace("xsi", "http://www.w3.org/2001/XMLSchema-instance");
  var noNS = XmlService.getNamespace("noNamespaceSchemaLocation", "amzn-envelope.xsd");
  
  var Header = XmlService.createElement('Header')
    .addContent(XmlService.createElement('DocumentVersion').setText("1.01"))
    .addContent(XmlService.createElement('MerchantIdentifier').setText('MERCHANT_IDENTIFIER'));
  var Envelope = XmlService.createElement('AmazonEnvelope')
    .setAttribute("noNamespaceSchemaLocation", "amzn-envelope.xsd", xsi)
    .addContent(Header)
    .addContent(XmlService.createElement('MessageType').setText('Price'));
  var products = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('MWS').getDataRange().getValues();

  for (i=1; i<products.length; i++) {
    var Price = XmlService.createElement('Price')
      .addContent(XmlService.createElement('SKU').setText(products[i][0]))
      .addContent(XmlService.createElement('StandardPrice').setAttribute("currency", "USD")
        .setText(products[i][4]));
    var Message = XmlService.createElement('Message')
      .addContent(XmlService.createElement('MessageID').setText(i))
      .addContent(Price);
    Envelope.addContent(Message);
  }
  var document = XmlService.createDocument(Envelope);
  var xml = XmlService.getPrettyFormat().format(document);
  Logger.log(xml)
  return ContentService.createTextOutput(xml);
}
