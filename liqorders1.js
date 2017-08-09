function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Automation Menu')
    .addItem('Liquidation Price Lookup', 'liqPriceSearch')
    .addItem('Blackwrap Price Lookup', 'blackPriceSearch')
    .addSeparator()
    .addItem('Import New Blackwrap', 'importBlackwrap')
    .addItem('Generate Price Estimates', 'generatePrices')
    .addItem('Import Price Estimates', 'importPrices')
    .addToUi();
}

function nono(sheet, itemCount) {
  /**
  * This function pulls information from the 'Ban List' sheet
  * and makes items returns if they are on said ban list.
  **/

  // Initialize Ban List sheet and get its size and values.
  var banSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Ban List');
  var banList = banSheet.getDataRange().getValues();
  var banCount = banSheet.getLastRow();
  
  // Store banned brands and ASINs in seperate arrays.
  var banBrand = [];
  var banASIN = [];
  for (var i = 1; i < banCount; i++) {
    if (banList[i][0] != "") {banBrand.push(banList[i][0]);}
    if (banList[i][1] != "") {banASIN.push(banList[i][1]);}
  }
  
  // Make items returns if they are on banned brands list.
  var items = sheet.getRange(6, 2, itemCount).getValues();
  for (var i = 0; i < itemCount; i++) {
    for (var j = 0; j < banBrand.length; j++) {
      if (items[i][0].indexOf(banBrand[j]) != -1) {
        sheet.getRange(6+i, 5).setValue('R');
      }
    }
  }

  // Make items returns if they are on banned ASIN list.
  var banASIN = ['B01IBF30M', 'B0MYVCXB0', 'B01I3BYYJK', 'B01LWWUEDR'];
  var itemASIN = sheet.getRange(6, 4, itemCount).getValues();
  for (var i = 0; i < itemCount; i++) {
    for (var j = 0; j < banASIN.length; j++) {
      if (itemASIN[i][0] == banASIN[j]) {
        sheet.getRange(6+i, 5).setValue('R');
      }
    }
  }
}

function liqPriceSearch() {
  /**
  * This script accomplishes the following:
  *   1. Copy formulas from Subtotal and My Buy Price
  *   2. Set spacing between first and second order to 3 blank rows
  *   3. Move UPC and Notes information
  *   4. Find and print sum of Count column
  *   5. Update VLOOKUP range
  *   6. Copy A/E/R and Amazon formulas
  *   7. Cache and paste values to overwrite the formulas
  **/
  
  // Initialize sheet.
  var sheet = SpreadsheetApp.getActiveSheet();
  
  // Find last item and cache its row number.
  var items = sheet.getDataRange.getValues();
  var lastSheetRow = sheet.getLastRow();
  var itemCount = 0;
  for (var i=5; i<lastSheetRow; i++) {
    itemCount++;
    if (items[i][1] == "") {
      break;
    }
  }
  Logger.log("Item Count: " + itemCount)
  
  // Copy formulas from Subtotal and My Buy Price.
  var subRange = sheet.getRange("A3:H4");
  subRange.copyTo(sheet.getRange(itemCount+5+1, 1, 2, 8));
  
  // Find and cache the number of blank rows between first and second order.
  var blankCount = 0;
  for (i=itemCount+6; i<2000; i++) {
    if (items[i][1] == "") {
      blankCount++;
    } else {
      break;
    }
  }
  Logger.log("Blank Count: " + blankCount)
  
  // Set spacing between first and second order to 5 blank rows.
  switch (true) {
    case (blankCount < 5):
      // Add blank rows.
      sheet.insertRowsAfter(itemCount+5+2, 5-blankCount);
      break;
    case (blankCount == 5):
      // Do nothing.
      break;
    case (blankCount > 5):
      // Delete rows.
      sheet.deleteRows(itemCount+5+3, blankCount-5);
      break;
    default:
      // Print error message.
      SpreadsheetApp.getUi().alert('Something went wrong counting blank rows.');
      return;
  }
  
  // Move UPC and Notes information.
  var upcRange = sheet.getRange(6, 6, itemCount);
  var notes = sheet.getRange(7, 7, itemCount-1).getValues();
  upcRange.copyTo(sheet.getRange(6, 4, itemCount));
  sheet.getRange(6, 4, itemCount).setNumberFormat('000000000000');
  sheet.getRange(7, 1, itemCount-1).setValues(notes);
  sheet.getRange(6, 5, itemCount).clear();
  
  // Find and print sum of Count and Amazon columns.
  var countA1 = sheet.getRange(6, 3, itemCount).getA1Notation();
  var amazonA1 = sheet.getRange(6, 7, itemCount).getA1Notation();
  var feesA1 = sheet.getRange(6, 8, itemCount).getA1Notation();
  sheet.getRange(itemCount+5+1, 3).setFormula("=SUM("+countA1+")");
  sheet.getRange(itemCount+5+1, 7).setFormula("=SUM("+amazonA1+")");
  sheet.getRange(itemCount+5+1, 8).setFormula("=SUM("+feesA1+")");
  
  // Create formula array.
  var formArray = [];
  for (var i=5; i<itemCount; i++) {
    var n = String(i+1);
    var L = String(5+itemCount);
    var b = String(lastSheetRow);
    formArray.push([
      '=IF($D'+n+'<>"",VLOOKUP($D'+n+',$D'+L+':$H'+b+',3), VLOOKUP($E'+n+',$E'+L+':$H'+b+',2))',
      '=IF($D'+n+'<>"",VLOOKUP($D'+n+',$D'+L+':$H'+b+',4), VLOOKUP($E'+n+',$E'+L+':$H'+b+',3))',
      '=IF($D'+n+'<>"",VLOOKUP($D'+n+',$D'+L+':$H'+b+',5), VLOOKUP($E'+n+',$E'+L+':$H'+b+',4))'
    ]);
  }
  
  // Update VLOOKUP range.
  sheet.getRange(6, 6, itemCount, 3).setFormulas(formArray);
  
  // Make banned items returns
  nono(sheet, itemCount);
  
  // Cache and paste values to overwrite the formulas.
  var vlookupValues = sheet.getRange(6, 5, itemCount, 3).getValues();
  sheet.getRange(6, 5, itemCount, 3).setValues(vlookupValues);
}

function blackPriceSearch() {
  /**
  * This script accomplishes the following:
  *   1. Copy formulas from Subtotal and My Buy Price
  *   2. Set spacing between first and second order to 3 blank rows
  *   3. Find and print sum of Count column
  *   4. Update VLOOKUP range
  *   5. Copy A/E/R and Amazon formulas
  *   6. Cache and paste values to overwrite the formulas
  **/
  
  // Initialize sheet.
  var sheet6 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('sheet6');
  
  // Find last item and cache its row number.
  var items = sheet6.getDataRange.getValues();
  var lastSheetRow = sheet6.getLastRow();
  var itemCount = 0;
  for (var i=5; i<lastSheetRow; i++) {
    itemCount++;
    if (items[i][1] == "") {
      break;
    }
  }
  Logger.log("Item Count: " + itemCount)
  
  // Find and cache the number of blank rows between first and second order.
  var blankCount = 0;
  for (i=itemCount+6; i<2000; i++) {
    if (items[i][1] == "") {
      blankCount++;
    } else {
      break;
    }
  }
  Logger.log("Blank Count: " + blankCount)
  
  // Set spacing between first and second order to 5 blank rows.
  switch (true) {
    case (blankCount < 5):
      // Add blank rows.
      sheet6.insertRowsAfter(itemCount+5+2, 5-blankCount);
      break;
    case (blankCount == 5):
      // Do nothing.
      break;
    case (blankCount > 5):
      // Delete rows.
      sheet6.deleteRows(itemCount+5+3, blankCount-5);
      break;
    default:
      // Print error message.
      SpreadsheetApp.getUi().alert('Something went wrong counting blank rows.');
      return;
  }
  
  // Set summation formulas.
  sheet6.getRange(itemCount+6, 1, 2, 9).setFontStyle('bold');
  sheet6.getRange(itemCount+6, 1).setValue("SUBTOTAL");
  sheet6.getRange(itemCount+7, 1).setValue("MY BUY PRICE");  
  var countA1 = sheet6.getRange(6, 3, itemCount).getA1Notation();
  var amazonA1 = sheet6.getRange(6, 7, itemCount).getA1Notation();
  var feesA1 = sheet6.getRange(6, 8, itemCount).getA1Notation();
  sheet6.getRange(itemCount+6, 3).setFormula("=SUM("+countA1+")");
  sheet6.getRange(itemCount+6, 7).setFormula("=SUM("+amazonA1+")");
  var feeSumA1 = sheet6.getRange(itemCount+6, 8).setFormula("=SUM("+feesA1+")").getA1Notation();
  var buyA1 = sheet6.getRange(itemCount+7, 8).setFormula("=SUM("+feeSumA1+"*0.6)").setBackgroundRGB(217, 234, 211).getA1Notation();
  sheet6.getRange(itemCount+6, 9, 2).setBackgroundRGB(255, 153, 0);
  sheet6.getRange(itemCount+7, 9).setFormula("=ROUND("+buyA1+"*0.92,2)");
  
  // Create formula array.
  var lastRow = sheet6.getLastRow();
  var rangeA1 = sheet6.getRange(itemCount+11, 4, lastRow-itemCount-10, 5).getA1Notation();
  var formArray = [];
  for (var i=6; i<itemCount+6; i++) {
    var asinA1 = sheet6.getRange(i, 4).getA1Notation();
    formArray.push([
      "=VLOOKUP("+asinA1+","+rangeA1+"3,FALSE)",
      "=VLOOKUP("+asinA1+","+rangeA1+"4,FALSE)",
      "=VLOOKUP("+asinA1+","+rangeA1+"5,FALSE)"
    ]);
  }
  
  // Update VLOOKUP range.
  sheet6.getRange(6, 6, itemCount, 3).setFormulas(formArray);
  
  // Make banned items returns
  nono(sheet6, itemCount);
  
  
  // Cache and paste values to overwrite the formulas.
  var vlookupValues = sheet6.getRange(6, 5, itemCount, 3).getValues();
  sheet6.getRange(6, 5, itemCount, 3).setValues(vlookupValues);
}

function importBlackwrap() {
  /**
  * This script imports a new blackwrap manifest into the Sheet6 tab.
  **/
  
  // Load active sheet and check to be sure it is a new blackwrap manifest.
  var blackwrapSheet = SpreadsheetApp.getActiveSheet();
  if (blackwrapSheet.getSheetName().length < 9) {
    SpreadsheetApp.getUi().alert('Please select the new blackwrap manifest before running script.');
    return;
  }
  
  // Load sheet6 and the values from the blackwrap manifest.
  var sheet6 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet6');
  var blackwrapValues = blackwrapSheet.getDataRange().getValues().slice(1);
  
  // Sort the new values alphabetically by their titles.
  blackwrapValues.sort(function(a, b) {
    if (a[12] === b[12]) {
      return 0;
    }
    else {
      return (a[12] < b[12]) ? -1 : 1;
    }
  });
  
  // Cache values to be transferred over to Sheet6.
  var titles = getCol(blackwrapValues, 12);
  var asins = getCol(blackwrapValues, 0);
  var LPNs = getCol(blackwrapValues, 1);
  
  // Create appropriate number of rows in Sheet6.
  var itemCount = blackwrapValues.length;
  sheet6.insertRowsBefore(6, blackwrapValues.length + 5);
  
  // Transfer values over to Sheet6.
  for (var i = 0; i < itemCount; i++) {
    var j = i + 6;
    sheet6.getRange(j, 2).setValue(titles[i]);
    sheet6.getRange(j, 3).setValue("1");
    sheet6.getRange(j, 4).setValue(asins[i]);
    sheet6.getRange(j, 5).setValue(LPNs[i]);
  }
  
  // Set summation formulas.
  sheet6.getRange(itemCount+6, 1, 2, 9).setFontStyle('bold');
  sheet6.getRange(itemCount+6, 1).setValue("SUBTOTAL");
  sheet6.getRange(itemCount+7, 1).setValue("MY BUY PRICE");  
  var countA1 = sheet6.getRange(6, 3, itemCount).getA1Notation();
  var amazonA1 = sheet6.getRange(6, 7, itemCount).getA1Notation();
  var feesA1 = sheet6.getRange(6, 8, itemCount).getA1Notation();
  sheet6.getRange(itemCount+6, 3).setFormula("=SUM("+countA1+")");
  sheet6.getRange(itemCount+6, 7).setFormula("=SUM("+amazonA1+")");
  var feeSumA1 = sheet6.getRange(itemCount+6, 8).setFormula("=SUM("+feesA1+")").getA1Notation();
  var buyA1 = sheet6.getRange(itemCount+7, 8).setFormula("=SUM("+feeSumA1+"*0.6)").setBackgroundRGB(217, 234, 211).getA1Notation();
  sheet6.getRange(itemCount+6, 9, 2).setBackgroundRGB(255, 153, 0);
  sheet6.getRange(itemCount+7, 9).setFormula("=ROUND("+buyA1+"*0.92,2)");
  
  // @TODO: blackwrap filter is not finding anything. FIX THIS.
  // Set order title.
//  var orders = getCol(blackwrapValues, 0);
//  var blackwraps = orders.filter(function(value) {
//    return value.substr(0, 5) == "BLACK";
//  });
//  var blackNum = blackwraps.length;
//  var n = -1;
//  for (var i = 0; i < blackwraps.length; i++) {
//    while (blackwraps[i].substr(n)[0] != "") {n--;}
//    if (parseInt(blackwraps[i].substr(n)) > blackNum) {
//      blackNum = blackwraps[i].substr(n);
//    }
//  }
//  sheet6.getRange(6, 1).setValue("BLACKWRAP "+String(blackNum+1));
  
  // Set VLOOKUP formulas.
  var lastRow = sheet6.getLastRow();
  var rangeA1 = sheet6.getRange(itemCount+11, 4, lastRow-itemCount-10, 5).getA1Notation();
  for (var i = 6; i < itemCount+6; i++) {
    var asinA1 = sheet6.getRange(i, 4).getA1Notation();
    sheet6.getRange(i, 6).setFormula("=VLOOKUP("+asinA1+","+rangeA1+",3,FALSE)");
    sheet6.getRange(i, 7).setFormula("=VLOOKUP("+asinA1+","+rangeA1+",4,FALSE)");
    sheet6.getRange(i, 8).setFormula("=VLOOKUP("+asinA1+","+rangeA1+",5,FALSE)");
  }
  
  // Make banned items returns
  nono(sheet6, itemCount);
  
  // Copy formula output values and paste them as text.
  var vlookupValues = sheet6.getRange(6, 6, itemCount, 3).getValues();
  sheet6.getRange(6, 6, itemCount, 3).setValues(vlookupValues);
}

function generatePrices() {
  /**
  * This script uses the Amazon Products API in tandem with
  * klasrun.com PHP scripting to create a JSON file of the
  * products in Sheet6 with their price, weight, and sales
  * rank on Amazon.
  */
  
  SpreadsheetApp.getUi().alert(
    'Go to the following URL and wait for a success message:\n\n'
    + 'http://klasrun.com/AmazonMWS/MarketplaceWebServiceProducts/Functions/BlackwrapPricing.php');
}

function importPrices() {
  /**
  * This script accomplishes the following tasks:
  *  1. Pull json file from MWS server
  *  2. Convert json into multidimensional array
  *  3. Update sheet6 with ASINs/UPCs.
  */
  
  // Fetch the json array from website and parse into JS object.
  var response = UrlFetchApp.fetch('http://klasrun.com/AmazonMWS/MarketplaceWebServiceProducts/Functions/blackwrap.json');
  var json = response.getContentText();
//  // Preserve newlines, etc - use valid JSON
//  json = json.replace(/\\n/g, "\\n")  
//               .replace(/\\'/g, "\\'")
//               .replace(/\\"/g, '\\"')
//               .replace(/\\&/g, "\\&")
//               .replace(/\\r/g, "\\r")
//               .replace(/\\t/g, "\\t")
//               .replace(/\\b/g, "\\b")
//               .replace(/\\f/g, "\\f");
//  // Remove non-printable and other non-valid JSON chars
//  json = json.replace(/[\u0000-\u0019]+/g,""); 
  var data = JSON.parse(json);
  
  // Convert data object into multidimensional array.
  // Ordering is same as in MWS tab.
  var itemCount = data.length;
  var itemArray = [];
  for (var i = 0; i < itemCount; i++) {
    var item = data[i];
    itemArray.push([
      // item.Title,
      // item.UPC,
      // item.ASIN,
      item.Price,
      item.Rank,
      item.Weight
    ]);
  }

  // Push array into sheet6 tab.
  var sheet6 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet6');
  sheet6.getRange(6, 10, itemCount, 3).setValues(itemArray);
}

function getCol(matrix, col){
// Take in a matrix and slice off a column from it.
// @param Col is 0-indexed.
  var column = [];
  var l = matrix.length;
  for(var i=0; i<l; i++){
     column.push(matrix[i][col]);
  }
  return column;
}
