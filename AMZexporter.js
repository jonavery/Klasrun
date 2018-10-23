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

function setComment(number) {
  switch (number) {
    case 1:
      return "Item will show signs of usage. Item may have mild to severe cosmetic and packaging damage. Item may be missing accessories. Item may not come in original packaging.";
    case 2:
      return "Item may show minor signs of wear and tear. Item includes all major components needed to function. Item may be missing non-essential accessories. Item may not come in original packaging.";
    case 3:
      return "Item may be missing original packaging and/or instruction manual. Item includes all major accessories and is fully functinal.";
    case 4:
      return "Item may be missing original packaging, or original packaging may be damaged. Item itself is in excellent, like new condition with no cosmetic damage or signs of use.";
    case 5:
      return "Item is brand new and has never been used. Item comes in unopened, original packaging.";
    default:
      return "";
  }
}

function getXML() {
  var root = XmlService.createElement('items');
  var items = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AMZ').getDataRange().getValues();
  for (i=1; i<items.length; i++) {
    if(!(items[i][8] == "1" || items[i][8] == "2" || items[i][8] == "3" || items[i][8] == "4" || items[i][8] == "5")) {continue;}
    if((items[i][17] == "E" || items[i][17] == "e")) {continue;}
    if((items[i][11] == "")) {continue;}
    var child = XmlService.createElement('item')
      .addContent(XmlService.createElement('SKU').setText(items[i][1]+items[i][16]))
      .addContent(XmlService.createElement('Title').setText(items[i][2]))
      .addContent(XmlService.createElement('ASIN').setText(items[i][3]))
      .addContent(XmlService.createElement('Condition').setText(getCondition(parseInt(items[i][8]))))
      .addContent(XmlService.createElement('Comment').setText(setComment(items[i][8])))
      .addContent(XmlService.createElement('Defect').setText(setMarkdown(parseInt(items[i][8]))));
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

