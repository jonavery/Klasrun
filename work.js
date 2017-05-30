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
    .addContent(XmlService.createElement('MessageType').setText('Product'))
    .addContent(XmlService.createElement('PurgeAndReplace').setText('false'));
  var products = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('MWS').getDataRange().getValues();

  for (i=1; i<products.length; i++) {
    if(products[i][4] == "" || products[i][4] == "undefined") {continue;}
    var StandardProductID = XmlService.createElement('StandardProductID')
      .addContent(XmlService.createElement('Type').setText('ASIN'))
      .addContent(XmlService.createElement('Value').setText(products[i][3]));
    var Condition = XmlService.createElement('Condition')
      .addContent(XmlService.createElement('ConditionType').setText(products[i][9]))
      .addContent(XmlService.createElement('ConditionNote').setText(products[i][10]));
    var Product = XmlService.createElement('Product')
      .addContent(XmlService.createElement('SKU').setText(products[i][0]))
      .addContent(StandardProductID)
      .addContent(Condition);
    var Message = XmlService.createElement('Message')
      .addContent(XmlService.createElement('MessageID').setText(i))
      .addContent(XmlService.createElement('OperationType').setText('Update'))
      .addContent(Product);
    Envelope.addContent(Message);
  }
  var document = XmlService.createDocument(Envelope);
  var xml = XmlService.getPrettyFormat().format(document);
  Logger.log(xml)
  return ContentService.createTextOutput(xml);
}
