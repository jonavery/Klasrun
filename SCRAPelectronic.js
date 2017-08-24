function doGet() {
  return getXML();
}

function getCondition(number) {
  switch (number) {
    case 0:
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

function setMarkdown(number) {
  switch (number) {
    case 0:
      return "yes";
    case 1:
    case 2:
    case 3:
    case 4:
    case 5:
    default:
      return "no";
  }
}

function getXML() {
  var root = XmlService.createElement('items');
  var items = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SCRAP').getDataRange().getValues();
  for (i=1; i<items.length; i++) {
    if(!(items[i][7] == "Y" || items[i][7] == "y") || items[i][11] == "") {continue;}
    if(!(items[i][17] == "E" || items[i][17] == "e")) {continue;}
    var child = XmlService.createElement('item')
      .addContent(XmlService.createElement('SKU').setText(items[i][1]))
      .addContent(XmlService.createElement('Title').setText(items[i][2]))
      .addContent(XmlService.createElement('ASIN').setText(items[i][3]))
      .addContent(XmlService.createElement('Condition').setText(getCondition(parseInt(items[i][16]))))
      .addContent(XmlService.createElement('Comment').setText(items[i][18]));
      .addContent(XmlService.createElement('Defect').setText(setMarkdown(parseInt(items[i][16]))));
    var grandchild = XmlService.createElement('Dimensions')
      .addContent(XmlService.createElement('Weight').setText(items[i][11]))
      .addContent(XmlService.createElement('Length').setText(items[i][12]))
      .addContent(XmlService.createElement('Width').setText(items[i][13]))
      .addContent(XmlService.createElement('Height').setText(items[i][14]));
    child.addContent(grandchild);
    root.addContent(child);
  }
  var document = XmlService.createDocument(root);
  var xml = XmlService.getPrettyFormat().format(document);
  Logger.log(xml)
  return ContentService.createTextOutput(xml);
}
