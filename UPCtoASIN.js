function doGet() {
  return getXML();
}

function getCondition(number) {
  switch (number) {
    case 1:
      return "UsedAcceptable";
    case 2:
      return "UsedGood";
    case 3:
      return "UsedVeryGood";
    case 4:
      return "UsedLikeNew";
    case 5:
      return "New";
    default:
      return number;
  }
}

function getXML() {
  var root = XmlService.createElement('items');
  var items = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SCRAP').getDataRange().getValues();
  for (i=1; i<items.length; i++) {
    if(!(items[i][6] == "Y" || items[i][6] == "y")) {continue;}
    var child = XmlService.createElement('item')
      .addContent(XmlService.createElement('SKU').setText(items[i][1]))
      .addContent(XmlService.createElement('Title').setText(items[i][2]))
      .addContent(XmlService.createElement('ASIN').setText(items[i][3]))
      .addContent(XmlService.createElement('Condition').setText(getCondition(parseInt(items[i][15]))))
      .addContent(XmlService.createElement('Comment').setText(items[i][17]));
    var grandchild = XmlService.createElement('Dimensions')
      .addContent(XmlService.createElement('Weight').setText(items[i][10]))
      .addContent(XmlService.createElement('Length').setText(items[i][11]))
      .addContent(XmlService.createElement('Width').setText(items[i][12]))
      .addContent(XmlService.createElement('Height').setText(items[i][13]));
    child.addContent(grandchild);
    root.addContent(child);
  }
  var document = XmlService.createDocument(root);
  var xml = XmlService.getPrettyFormat().format(document);
  Logger.log(xml)
  return ContentService.createTextOutput(xml);
}
