function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Automation Menu')
    .addItem('Import New Blackwrap', 'importBlackwrap')
    .addItem('Generate Price Estimates', 'generatePrices')
    .addItem('Import Price Estimates', 'importPrices')
    .addSeparator()
    .addItem('Update Export', 'updateLiqFormat')
    .addItem('Export to LIQ & WORK', 'exportData')
    .addSeparator()
    .addItem('Update ASINs', 'updateASINs')
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
        sheet.getRange(6+i, 6).setValue('R');
        sheet.getRange(6+i, 1).setValue('BAN');
      }
    }
  }

  // Make items returns if they are on banned ASIN list.
  var banASIN = ['B01IBF30M', 'B0MYVCXB0', 'B01I3BYYJK', 'B01LWWUEDR'];
  var itemASIN = sheet.getRange(6, 4, itemCount).getValues();
  for (var i = 0; i < itemCount; i++) {
    for (var j = 0; j < banASIN.length; j++) {
      if (itemASIN[i][0] == banASIN[j]) {
        sheet.getRange(6+i, 6).setValue('R');
        sheet.getRange(6+i, 1).setValue('BAN');
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
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Research');
  
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
  
  // Set summation formulas.
  sheet.getRange(itemCount+6, 1, 2, 9).setFontStyle('bold');
  sheet.getRange(itemCount+6, 1).setValue("SUBTOTAL");
  sheet.getRange(itemCount+7, 1).setValue("MY BUY PRICE");  
  var countA1 = sheet.getRange(6, 3, itemCount).getA1Notation();
  var amazonA1 = sheet.getRange(6, 7, itemCount).getA1Notation();
  var feesA1 = sheet.getRange(6, 8, itemCount).getA1Notation();
  sheet.getRange(itemCount+6, 3).setFormula("=SUM("+countA1+")");
  sheet.getRange(itemCount+6, 7).setFormula("=SUM("+amazonA1+")");
  var feeSumA1 = sheet.getRange(itemCount+6, 8).setFormula("=SUM("+feesA1+")").getA1Notation();
  var buyA1 = sheet.getRange(itemCount+7, 8).setFormula("=SUM("+feeSumA1+"*0.6)").setBackgroundRGB(217, 234, 211).getA1Notation();
  sheet.getRange(itemCount+6, 9, 2).setBackgroundRGB(255, 153, 0);
  sheet.getRange(itemCount+7, 9).setFormula("=ROUND("+buyA1+"*0.92,2)");
  
  // Create formula array.
  var lastRow = sheet.getLastRow();
  var rangeA1 = sheet.getRange(itemCount+11, 4, lastRow-itemCount-10, 5).getA1Notation();
  var formArray = [];
  for (var i=6; i<itemCount+6; i++) {
    var asinA1 = sheet.getRange(i, 4).getA1Notation();
    formArray.push([
      "=VLOOKUP("+asinA1+","+rangeA1+"3,FALSE)",
      "=VLOOKUP("+asinA1+","+rangeA1+"4,FALSE)",
      "=VLOOKUP("+asinA1+","+rangeA1+"5,FALSE)"
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

function importBlackwrap() {
  /**
  * This script imports a new blackwrap manifest into the sheet tab.
  **/
  
  // Load active sheet and check to be sure it is a new blackwrap manifest.
  var blackwrapSheet = SpreadsheetApp.getActiveSheet();
  if (blackwrapSheet.getSheetName().length < 9) {
    SpreadsheetApp.getUi().alert('Please select the new blackwrap manifest before running script.');
    return;
  }
  
  // Load sheet and the values from the blackwrap manifest.
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Research');
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
  
  // Cache values to be transferred over to sheet.
  for (var i = 0; i < blackwrapSheet.getLastColumn(); i++) {                                                                                                  
    var header = blackwrapSheet.getRange(1, i+1).getValue();                                                                                                  
    if (header == "Item Description") {var titleCol = i;}                                                                                                     
    if (header == "B00 ASIN") {var asinCol = i;}                                                                                                              
    if (header == "X-Z ASIN") {var lpnCol = i;}                                                                                                               
  }                                                                                                                                                           
  var titles = getCol(blackwrapValues, titleCol);                                                                                                             
  var asins = getCol(blackwrapValues, asinCol);                                                                                                               
  var LPNs = getCol(blackwrapValues, lpnCol); 
  
  // Create appropriate number of rows in sheet.
  var itemCount = blackwrapValues.length;
  sheet.insertRowsBefore(6, blackwrapValues.length + 5);
  
  // Transfer values over to sheet.
  for (var i = 0; i < itemCount; i++) {
    var j = i + 6;
    sheet.getRange(j, 2).setValue(titles[i]);
    sheet.getRange(j, 3).setValue("1");
    sheet.getRange(j, 4).setValue(asins[i]);
    sheet.getRange(j, 5).setValue(LPNs[i]);
  }
  
  // Set summation formulas.
  sheet.getRange(itemCount+6, 1, 2, 9).setFontStyle('bold');
  sheet.getRange(itemCount+6, 1).setValue("SUBTOTAL");
  sheet.getRange(itemCount+7, 1).setValue("MY BUY PRICE");  
  var countA1 = sheet.getRange(6, 3, itemCount).getA1Notation();
  var amazonA1 = sheet.getRange(6, 7, itemCount).getA1Notation();
  var feesA1 = sheet.getRange(6, 8, itemCount).getA1Notation();
  sheet.getRange(itemCount+6, 3).setFormula("=SUM("+countA1+")");
  sheet.getRange(itemCount+6, 7).setFormula("=SUM("+amazonA1+")");
  var feeSumA1 = sheet.getRange(itemCount+6, 8).setFormula("=SUM("+feesA1+")").getA1Notation();
  var buyA1 = sheet.getRange(itemCount+7, 8).setFormula("=SUM("+feeSumA1+"*0.6)").setBackgroundRGB(217, 234, 211).getA1Notation();
  sheet.getRange(itemCount+6, 9, 2).setBackgroundRGB(255, 153, 0);
  sheet.getRange(itemCount+7, 9).setFormula("=ROUND("+buyA1+"*0.92,2)");
  
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
//  sheet.getRange(6, 1).setValue("BLACKWRAP "+String(blackNum+1));
  
  // Set VLOOKUP formulas.
  var lastRow = sheet.getLastRow();
  var rangeA1 = sheet.getRange(itemCount+11, 4, lastRow-itemCount-10, 5).getA1Notation();
  for (var i = 6; i < itemCount+6; i++) {
    var asinA1 = sheet.getRange(i, 4).getA1Notation();
    sheet.getRange(i, 6).setFormula("=VLOOKUP("+asinA1+","+rangeA1+",3,FALSE)");
    sheet.getRange(i, 7).setFormula("=VLOOKUP("+asinA1+","+rangeA1+",4,FALSE)");
    sheet.getRange(i, 8).setFormula("=VLOOKUP("+asinA1+","+rangeA1+",5,FALSE)");
  }
  
  // Make banned items returns
  nono(sheet, itemCount);
  
  // Copy formula output values and paste them as text.
  var vlookupValues = sheet.getRange(6, 6, itemCount, 3).getValues();
  sheet.getRange(6, 6, itemCount, 3).setValues(vlookupValues);
}

function generatePrices() {
  /**
  * This script uses the Amazon Products API in tandem with
  * klasrun.com PHP scripting to create a JSON file of the
  * products in sheet with their price, weight, and sales
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
  *  3. Update sheet with ASINs/UPCs.
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

  // Push array into sheet tab.
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Research');
  sheet.getRange(6, 10, itemCount, 3).setValues(itemArray);
}

function updateLiqFormat() {
  /************************************************************************
  * This script accomplishes the following tasks:
  *   1. Find order in Research from its number in Cycles
  *   2. Move order information into Export with correct formatting
  *   3. Fill out all relevant formulas on the right side of Export
  *   4. Adjust per item cost to align with total cost
  *************************************************************************/
  
  // Set ID for the spreadsheet file to be used.
  var maniID = "1TaxBUL8WjTvV3DjJEMduPK6Qs3A5GoFDmZHiUcc-LUY";
  
  // Initialize the sheets to be accessed.
  var sheetExp = SpreadsheetApp.openById(maniID).getSheetByName("Export");
  var sheetResearch = SpreadsheetApp.openById(maniID).getSheetByName("Research");
  var sheetCycles = SpreadsheetApp.openById(maniID).getSheetByName("Cycles");
  
  // Extract first column from Research sheet.
  var orders = sheetResearch.getDataRange().getValues();
  var orderCol = (getCol(orders,0));
  
  var auctions = sheetCycles.getDataRange().getValues();
  
  // Save today's properly formatted date as a variable.
  var today = new Date();
  var dd = today.getDate();
  var mm = today.getMonth()+1; // .getMonth is 0-indexed.
  var yyyy = today.getFullYear();
  if(dd<10) { dd = '0' + dd;}
  if(mm<10) { mm = '0' + mm;}
  var today = mm + '/' + dd + '/' + yyyy;
  
  
  // Initialize counting variable for reporting number of items copied.
  var itemCopy = 0;
  
  // Clear Export sheet and remove empty rows.
  sheetExp.getRange(3, 1, sheetExp.getMaxRows()-2, sheetExp.getLastColumn()).clear();
  sheetExp.deleteRows(3, sheetExp.getMaxRows()-2);
  
  // Cache order ID and item count, then insert appropriate number of rows into Export sheet.
  var orderID = auctions[3][10];
  var itemCount = auctions[3][13];
  sheetExp.insertRowsAfter(2, itemCount);
  // Find order in Research sheet
  var orderIndex = orderCol.indexOf(orderID);
  if (orderIndex == -1) {
    SpreadsheetApp.getUi().alert('Could not find order #' + orderID + '. Aborting...');
    return;
  }
  // Save order as range with itemCount number of rows
  var orderItems = sheetResearch.getRange(orderIndex+1, 2, itemCount, 5);
  // Copy range values over to Export
  orderItems.copyValuesToRange(sheetExp, 3, 7, 3, 3+itemCount);
  sheetExp.getRange(3, 1, itemCount, 9).setBackground('white');
    
  // Copy A/E/R from Buy Site to correct column
  var AER = sheetExp.getRange(3, 8, itemCount);
  sheetExp.getRange(3, 7, itemCount).moveTo(AER);
  var formulaRange = sheetExp.getRange(2, 10, 1, 5);
  // Fill in date, buy site, and cost information.
  for (var j=1; j <= itemCount; j++) {
    sheetExp.getRange(2+j, 2).setValue(today);
    sheetExp.getRange(2+j, 7).setValue("LIQUIDATION");
    sheetExp.getRange(2+j, 9).setValue(orderID);
    formulaRange.copyTo(sheetExp.getRange(2+j, 10, 1, 5));
  }
  // Copy per item cost values.
  sheetExp.getRange(3, 15, itemCount).setValue(sheetExp.getRange(3, 14, itemCount).getDisplayValues());
  // Compare rounded cost to actual cost
  var prices = sheetExp.getRange(3, 15, itemCount).getValues();
  var orderTotal = sheetExp.getRange(3,11).getValue();
  var roundedTotal = Number(round(sumArray(prices), 2));
  if (roundedTotal < orderTotal) {
    // If rounded is lower, compensate top per item cost
    sheetExp.getRange(3, 15).setValue(Number(prices[0]) + orderTotal - roundedTotal);
  }
  else if (roundedTotal > orderTotal) {
    // If rounded is higher, compensate bottom per item cost
    sheetExp.getRange(2+itemCount, 15).setValue(Number(prices[itemCount-1]) + orderTotal - roundedTotal);
  }
  itemCopy = itemCopy + itemCount;
  // Post dialogue box showing # of orders and items copied to LIQ FORMAT.
  SpreadsheetApp.getUi().alert('Script finished.\n\nItems Copied: ' + itemCopy);
}

function exportData() {
  /*********************************************************************************************
  * This function will accomplish the following:
  *   1. Load the relevant Liquidation, Work, and Manifest sheets.
  *   2. Save the needed ranges from the Manifest sheet.
  *   3. Copy over the ranges into the Liquidation and Work sheets with the proper formatting.
  *   4. Fill in any constant values in the Liquidation and Work Sheets.
  **********************************************************************************************/
  
  // Set ID's for the spreadsheet file to be used.
  var maniID = "1TaxBUL8WjTvV3DjJEMduPK6Qs3A5GoFDmZHiUcc-LUY";
  var workID = "1okDFF9236lGc4vU6W7HOD8D-3ak8e_zntehvFatYxnI";
  var liqID = "1Xqsc6Qe_hxrWN8wRd_vgdBdrCtJXVlvVC9w53XJ0BUM";
  
  // Load the sheets between which data will be transferred.
  var sheetExport = SpreadsheetApp.openById(maniID).getSheetByName("LIQ FORMAT");
  var sheetFuture = SpreadsheetApp.openById(workID).getSheetByName("Future Listing");
  var sheetLiquid = SpreadsheetApp.openById(liqID).getSheetByName("Liquidation Orders");
  
  // Save last row in each sheet to be used for indexing later.
  var maniLastRow = sheetExport.getLastRow();
  var liqLastRow = sheetLiquid.getLastRow();
  
  // Load all of the values from manifest sheet.
  var maniValues = sheetExport.getDataRange().getValues();
  
  // Prepare the future listings sheet for data entry.
  var futureMaxRows = sheetFuture.getMaxRows();
  sheetFuture.getRange(2, 1, futureMaxRows-1, sheetFuture.getLastColumn()).clear();
  var futureNeededRows = maniValues.length + 1 - futureMaxRows;
  switch(true) {
    case (futureNeededRows > 0):
      // Add blank rows.
      sheetFuture.insertRowsAfter(1, futureNeededRows);
      break;
    case (futureNeededRows == 0):
      // Do nothing.
      break;
    case (futureNeededRows < 0):
      // Delete rows.
      sheetFuture.deleteRows(2, -futureNeededRows);
      break;
    default:
      // Print error message.
      SpreadsheetApp.getUi().alert('Something went wrong formatting blank rows.');
      return;
  }
  
  // Add the proper number of rows to the liquidation sheet if needed.
  var liqNeededRows = maniValues.length + liqLastRow - sheetLiquid.getMaxRows();
  switch(true) {
    case (liqNeededRows > 0):
      // Add blank rows.
      sheetLiquid.insertRowsAfter(liqLastRow, liqNeededRows);
      break;
    case (liqNeededRows == 0):
    case (liqNeededRows < 0):
      // Do nothing.
      break;
    default:
      // Print error message.
      SpreadsheetApp.getUi().alert('Something went wrong formatting blank rows.');
      return;
  }
  
  // Load highest SKU # from liquidation sheet.
  var allSKUs = getCol(sheetLiquid.getRange(2, 1, liqLastRow-1).getValues(), 0);
  var highSKU = 1;
  for (i=0; i < allSKUs.length; i++) {
    if (allSKUs[i] > highSKU) {
      var highSKU = allSKUs[i];
    }
  }
  
  // Cache all order numbers currently in liquidation sheet and check to see if data has already been transferred.
  var allOrderNums = getCol(sheetLiquid.getRange(1, 8, liqLastRow).getValues(), 0);
  for (var i=2; i < maniLastRow; i++) {   
    var k = i-1;
    if (allOrderNums.indexOf(maniValues[i][8]) > -1) {
      Logger.log('Order #' + maniValues[i][8] + ' has already been copied.');
      break;
    }
    // To Future(column): Title(3), ASIN(4), LPN(5), A/E/R(6), and 7-digit Order #(7) from liq orders.
    sheetFuture.getRange(i, 2).setValue(highSKU + k);       // SKU
    sheetFuture.getRange(i, 3).setValue(maniValues[i][2]);  // Title
    sheetFuture.getRange(i, 4).setValue(maniValues[i][4]);  // ASIN
    sheetFuture.getRange(i, 5).setValue(maniValues[i][5]);  // LPN
    sheetFuture.getRange(i, 6).setValue(maniValues[i][7]);  // A/E/R
    sheetFuture.getRange(i, 7).setValue(maniValues[i][9]);  // Order #
    // To Liquid(column): Date(2), Title(3), Quantity(4), ASIN(5), Buy Site(6), A/E/R(7), 7-digit #(8), Buy Price(11), and Card(12) from liq orders.
    sheetLiquid.getRange(liqLastRow + k, 1).setValue(highSKU + k);       // SKU
    sheetLiquid.getRange(liqLastRow + k, 2).setValue(maniValues[i][1]);  // Date
    sheetLiquid.getRange(liqLastRow + k, 3).setValue(maniValues[i][2]);  // Title
    sheetLiquid.getRange(liqLastRow + k, 4).setValue(maniValues[i][3]);  // Quantity
    sheetLiquid.getRange(liqLastRow + k, 5).setValue(maniValues[i][4]);  // ASIN
    sheetLiquid.getRange(liqLastRow + k, 6).setValue(maniValues[i][6]);  // Buy Site
    sheetLiquid.getRange(liqLastRow + k, 7).setValue(maniValues[i][7]);  // A/E/R
    sheetLiquid.getRange(liqLastRow + k, 8).setValue(maniValues[i][9]);  // Order #
    sheetLiquid.getRange(liqLastRow + k, 9).setValue("FBA");             // Sell Site
    sheetLiquid.getRange(liqLastRow + k, 10).setValue("FBA");            // Sell Order
    sheetLiquid.getRange(liqLastRow + k, 11).setValue(maniValues[i][14]);// Buy Price
    sheetLiquid.getRange(liqLastRow + k, 12).setValue(maniValues[i][11]);// Card
    // Setup liquidation formulas for new entry.
    var r = String(liqLastRow + k);
    sheetLiquid.getRange(liqLastRow + k, 14).setFormula("=M"+r+"-K"+r);  // Actual Profit
    sheetLiquid.getRange(liqLastRow + k, 15).setFormula("=M"+r+"/K"+r);  // Actual % Increase
    sheetLiquid.getRange(liqLastRow + k, 22).setFormula("=VLOOKUP(A"+r+",Returns!A:A,1,0)");        // RETURNS V
    sheetLiquid.getRange(liqLastRow + k, 23).setFormula("=VLOOKUP(A"+r+",Salvage!A:A,1,0)");        // SALVAGE V
    sheetLiquid.getRange(liqLastRow + k, 24).setFormula("=VLOOKUP(A"+r+",Reimbursements!F:F,1,0)"); // REIMBURSE V
    sheetLiquid.getRange(liqLastRow + k, 25).setFormula("=VLOOKUP(A"+r+",Inventory!B:B,1,0)");      // INVENTORY V
//    sheetLiquid.getRange(liqLastRow + k, 26).setFormula("=VLOOKUP(A"+r+",Connor!G:H,2,0)");         // FBA SHIPMENT STATUS
//    sheetLiquid.getRange(liqLastRow + k, 27).setFormula("=VLOOKUP(A"+r+",Connor!K:K,1,0)");         // FBA SHIPMENT ISSUE
  }
}

function updateASINs() {
  /************************************************************************
  * This script accomplishes the following tasks:
  *   1. Find unique items in Research
  *   2. Filter out duplicate items and unresearched items
  *   3. Move remaining items into ASINs
  *************************************************************************/
  
  // Set ID for the spreadsheet file to be used.
  var maniID = "1TaxBUL8WjTvV3DjJEMduPK6Qs3A5GoFDmZHiUcc-LUY";
  
  // Initialize the sheets to be accessed.
  var sheetASINs = SpreadsheetApp.openById(maniID).getSheetByName("ASINs");
  var sheetResearch = SpreadsheetApp.openById(maniID).getSheetByName("Research");

  // Cache and count ASINs from Research and ASINs.
  var researchValues = sheetResearch.getDataRange().getValues();
  var asinsValues = sheetASINs.getDataRange().getValues();
  
  // Compare Research items to Database items.
    // Move item into Database if new, non-empty, and researched.
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

function rep(obj, n) {
  // Create function that makes an array of n length and obj identical inputs.
  var arr = [[]];
  for (i=0; i < n; i++) {arr[i][0].push(obj);}
  return arr;
}

function sumArray(array) {
  // Create function that finds the sum of an array.
  for (
    var
    i = 0,
    length = array.length,      // Cache the array length
    sum = 0;                    // The total amount
    i < length;                 // The "for"-loop condition
    sum += Number(array[i++])   // Add number on each iteration
  );
  return sum;
}

function containedIn(needles, haystack) {
  // Create function that checks to see if objects from one array are contained in a second array.
  var check = [];
  for (i=0; i < needles.length; i++) {
    check[i] = haystack.indexOf(needles[i][0]) > -1;
  }
  return check.indexOf(true) > -1;
}

function round(value, exp) {
  // Create function that rounds a value to exp decimal places
  if (typeof exp === 'undefined' || +exp === 0)
    return Math.round(value);
  value = +value;
  exp = +exp;
  if (isNaN(value) || !(typeof exp === 'number' && exp % 1 === 0))
    return NaN;
  // Shift
  value = value.toString().split('e');
  value = Math.round(+(value[0] + 'e' + (value[1] ? (+value[1] + exp) : exp)));
  // Shift back
  value = value.toString().split('e');
  return +(value[0] + 'e' + (value[1] ? (+value[1] - exp) : -exp));
}

