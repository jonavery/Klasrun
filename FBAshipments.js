function populateMWS() {
  /**
  * This script accomplishes the following tasks:
  *  1. Pull json file from MWS server
  *  2. Convert json into multidimensional array
  *  3. Push array into MWS tab.
  */
   
  // Fetch the json array from website and parse into JS object.
  var response = UrlFetchApp.fetch('http://klasrun.com/AmazonMWS/FBAInboundServiceMWS/Functions/FBA.json');
  var json = response.getContentText();
  var data = JSON.parse(json);
   
  // Convert data object into multidimensional array.
  // Ordering is same as in MWS tab.
  var itemCount = data.length;
  var itemArray = makeArray(8, itemCount, "");
  for (i = 0; i < itemCount; i++) {
    var item = data[i];
    itemArray[i] = ([
      item.Status,
      item.ShipmentId,
      item.SellerSKU,
      item.Status,
      item.Created,
      item.Updated,
      item.QuantityShipped,
      item.QuantityReceived
    ]);
  }
 
  // Initialize spreadsheet.
  var sheetFBA = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('1000');
  var lastRow = sheetFBA.getLastRow();
  
  // Add the proper number of rows to the liquidation sheet if needed.
  var neededRows = itemCount + lastRow - sheetFBA.getMaxRows();
  switch(true) {
    case (neededRows > 0):
      // Add blank rows.
      sheetFBA.insertRowsAfter(lastRow, neededRows);
      break;
    case (neededRows == 0):
    case (neededRows < 0):
      // Do nothing.
      break;
    default:
      // Print error message.
      SpreadsheetApp.getUi().alert('Something went wrong formatting blank rows.');
      return;
  }
  
  // Push array into FBA sheet.
  sheetFBA.getRange(2, lastRow+1, itemCount, 8).setValues(itemArray);
  
  // Cache month and year.
  var monthNames = ["January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"];
  var d = new Date();
  var month = monthNames[d.getMonth()].toUpperCase();
  var year = d.getFullYear();
  var date = month.slice(0,3) + ". '" + year.slice(2);
  
  // Set formulas for all new rows.
  for (var i = 0; i < itemCount; i++) {
    var r = i +lastRow + 1;
    var rangeFormula = sheetFBA.getRange(r, 9, 1, 2);
    var formulas = ([
        "=IFERROR(H" + r + "/G" + r + ",0)",
        "=IF(ISBLANK(I" + r + "),'',IF(I" + r + "=1,'OK',IF(I" + r + ">1,'OK: EXTRA',E" + r + "+45)))"
      ]);
    sheetFBA.getRange(r, 11).setValue(date);
    sheetFBA.getRange(r, 14).setValue(itemArray[i][1]);
  }
}
