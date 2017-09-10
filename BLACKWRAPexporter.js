function doGet() {
  return getXML();
}

function getXML() {
  var items = XmlService.createElement('items');
  var itemValues = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Research').getDataRange().getValues();
  var length = itemValues.length;
  
  for (var i = 5; i < length; i++) {
    if(itemValues[i][1] == "") {break;}
    var item = XmlService.createElement('item')
      .addContent(XmlService.createElement('Title').setText(itemValues[i][1]))
      .addContent(XmlService.createElement('ASIN').setText(itemValues[i][3]))
      .addContent(XmlService.createElement('UPC').setText(itemValues[i][4]))
    items.addContent(item);
  }
  
  var document = XmlService.createDocument(items);
  var xml = XmlService.getPrettyFormat().format(document);
  Logger.log(xml)
  return ContentService.createTextOutput(xml);
}

