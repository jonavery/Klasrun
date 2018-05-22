function doGet() {
  return getXML();
}

function getXML() {
  var Items = XmlService.createElement('items');
  var MWSsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('MWS');
  var products = MWSsheet.getDataRange().getValues();
  var lastRow = MWSsheet.getLastRow();
  var range = MWSsheet.getRange(2, 1, lastRow-1, 12);
  range.sort([{column: 1, ascending: true}, {column: 12, ascending: true}]);

  for (var i = 1; i < products.length; i++) {
    if(products[i][4] == "" || products[i][4] == "undefined") {continue;}
    var Dimensions = XmlService.createElement('Dimensions')
      .addContent(XmlService.createElement('Weight').setText(products[i][5]))
      .addContent(XmlService.createElement('Length').setText(products[i][6]))
      .addContent(XmlService.createElement('Width').setText(products[i][7]))
      .addContent(XmlService.createElement('Height').setText(products[i][8]));
    var Member = XmlService.createElement('Member')
      .addContent(XmlService.createElement('SellerSKU').setText(products[i][0]))
      .addContent(XmlService.createElement('Quantity').setText('1'))
      .addContent(XmlService.createElement('ShipmentId').setText(products[i][11]))
      .addContent(Dimensions);
    Items.addContent(Member);
  }

  var document = XmlService.createDocument(Items);
  var xml = XmlService.getPrettyFormat().format(document);
  Logger.log(xml)
  return ContentService.createTextOutput(xml);
}

