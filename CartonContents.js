function doGet() {
  return getXML();
}

function makeArray(w, h, val) {
// Create array with 'w' columns, 'h' rows, and filled with 'val'
  var arr = [];
  for(i = 0; i < h; i++) {
    arr[i] = [];
    for(j = 0; j < w; j++) {
      arr[i][j] = val;
    }
  }
  return arr;
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
    
  // Fetch the json array from website and parse into JS object.
  var response = UrlFetchApp.fetch('http://klasrun.com/AmazonMWS/FBAInboundServiceMWS/Functions/shipID.json');
  var json = response.getContentText();
  var data = JSON.parse(json);
   
  // Convert data object into multidimensional array.
  var shipments = Object.keys(data);
  var shipCount = shipments.length;
  var itemArray = {};
  var itemCount = [];
  var k = 0;
  for (var i = 0; i < shipCount; i++) {
    var id = shipments[i];
    itemCount[i] = data[id].length;
    
    for (var j = 0; j < itemCount[i]; j++) {
      itemArray[data[id][j]] = id;
    }
  }
 
  // Initialize sheet variables.
  var sheetMWS = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('MWS');
  var lastRow = sheetMWS.getLastRow();
  var SKUs = sheetMWS.getRange(1, 1, lastRow).getValues();
  
  // Import shipmentId's into sheet.
  sheetMWS.getRange(2, 12, lastRow-1, 1).clearContent();
  for (var i = 1; i < lastRow; i++) {
    sheetMWS.getRange(i+1, 12).setValue(itemArray[SKUs[i][0]]);
  }
  
  // Sort MWS sheet by shipmentId.
  var products = sheetMWS.sort(12).getDataRange().getValues();
  
  // l = index of current item within products array
  // k = index of shipment in counts array
  // j = cartonid number
  var l = 1;
  for (var k = 0; k < shipCount; k++) {
    var ShipmentId = shipments[k];
  
    var CartonContentsRequest = XmlService.createElement('CartonContentsRequest')
      .addContent(XmlService.createElement('ShipmentId').setText(ShipmentId))
      .addContent(XmlService.createElement('NumCartons').setText(itemCount[k]));
  
    for (var i = 0; i < itemCount[k]; i++) {
      var Item = XmlService.createElement('Item')
        .addContent(XmlService.createElement('SKU').setText(products[l][0]))
        .addContent(XmlService.createElement('QuantityShipped').setText('1'))
        .addContent(XmlService.createElement('QuantityInCase').setText('1'));
      var Carton = XmlService.createElement('Carton')
        .addContent(XmlService.createElement('CartonId').setText(products[l][0]))
        .addContent(Item);
      CartonContentsRequest.addContent(Carton);
      l++;
    }

    var Message = XmlService.createElement('Message')
      .addContent(XmlService.createElement('MessageID').setText(String(k+1)))
      .addContent(CartonContentsRequest);
    Envelope.addContent(Message);
  }
  
  var document = XmlService.createDocument(Envelope);
  var xml = XmlService.getPrettyFormat().format(document);
  Logger.log(xml)
  return ContentService.createTextOutput(xml);
}
