function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Automation Menu')
    .addItem('Import New Auctions', 'importLiqOrders')
    .addItem('Update LIQ FORMAT', 'updateLiqFormat')
    .addSeparator()
    .addItem('Export to LIQ & WORK', 'exportData')
    .addToUi();
}

function getCol(matrix, col){
  // Create function that strips specified column from an array.
  var column = [];
  var l = matrix.length;
  for(var i=0; i<l; i++){
     column.push(matrix[i][col]);
  }
  return column;
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

function importLiqOrders() {
  /*************************************************************************
  * This script accomplishes the following:
  *   1. Import the new auctions for Liq Orders1.ods
  *   2. Clear the LIQ FORMAT, Orders, and Auctions sheets from Manifest.
  *   3. Paste the auctions into Orders and Auctions.
  **************************************************************************/
  
  // Set ID's for the spreadsheet files to be used.
  var maniID = "1aV0bturINsUIgF-X8LcWZaWfehWEkhr6nUIVhbqdAxM";
  var liqOrdersID = "1TaxBUL8WjTvV3DjJEMduPK6Qs3A5GoFDmZHiUcc-LUY";
  var liquidID = "1Xqsc6Qe_hxrWN8wRd_vgdBdrCtJXVlvVC9w53XJ0BUM";
  
  // Initialize the sheets to be updated.
  var sheetLiq = SpreadsheetApp.openById(maniID).getSheetByName("LIQ FORMAT");
  var sheetOrders = SpreadsheetApp.openById(maniID).getSheetByName("Copy of Liq Orders Scrap");
  var sheetOldAuctions = SpreadsheetApp.openById(maniID).getSheetByName("Auctions");
  
  // Initialize liq orders1.ods sheets.
  var sheetNewAuctions = SpreadsheetApp.openById(liqOrdersID).getSheetByName("AUCTION");
  var sheetNewOrders = SpreadsheetApp.openById(liqOrdersID).getSheetByName("Sheet6")
  
  // Initialize liquidation sheet and cache current manifest sheet values.
  var sheetLiquid = SpreadsheetApp.openById(liquidID).getSheetByName("Liquidation Orders");
  var maniValues = sheetLiq.getDataRange().getValues();
  
  // Check to see if data has already been transferred.
  if (sheetLiq.getLastRow() > 2) {
    // Cache all order numbers currently in liquidation sheet.
    var currentOrderNums = getCol(sheetLiquid.getRange(1, 8, sheetLiquid.getLastRow()).getValues(), 0);
    if (currentOrderNums.indexOf(maniValues[2][8]) < 0 && maniValues[2][8] != "#N/A") {
      SpreadsheetApp.getUi().alert('LIQ FORMAT has not been transferred to LIQ and WORK yet. Transfer data before updating auctions.');
      return;
    }
  }
  
  // Prompt user for SKU and total number of auctions.
  var ui = SpreadsheetApp.getUi();
  var response1 = ui.prompt('Auctions on right(R) or left(L)?');
  var side = response1.getResponseText();
  var response2 = ui.prompt('Total number of auctions?');
  var numAuctions = response2.getResponseText();
  switch (side) {
    case 'r':
    case 'R':
    case 'right':
      var rangeAuctions = sheetNewAuctions.getRange(3, 11, sheetNewAuctions.getLastRow()-2, 6);
      break;
    case 'l':
    case 'L':
    case 'left':
      var rangeAuctions = sheetNewAuctions.getRange(3, 2, sheetNewAuctions.getLastRow()-2, 6);
      break;
    default:
      ui.alert('Invalid side entry. Assuming right side.');
      var rangeAuctions = sheetNewAuctions.getRange(3, 2, sheetNewAuctions.getLastRow()-2, 6);
      break;
  }
  
  // Clean out empty rows from new auctions.
  var newAuctions = rangeAuctions.getValues();
  newAuctions = newAuctions.sort();
  while (newAuctions[0][0] == "") {
    newAuctions = newAuctions.slice(1);
  }
  newAuctions = newAuctions.slice(0, numAuctions);
  
  // Loop through all old auctions and compare them to new auctions.
  if (sheetOldAuctions.getLastRow() > 0){
    var oldAuctions = getCol(sheetOldAuctions.getRange(1,1,sheetOldAuctions.getLastRow()).getValues(), 0);
    for (var i=0; i < oldAuctions.length; i++) {
      for(var j=0; j < newAuctions.length; j++) {
        if (oldAuctions[i] == newAuctions[j][0]) {
          newAuctions.splice(j, 1);
          Logger.log('Old i: ' + i + '    New j: ' + j);
        }
      }
    }
  }
      
  // If no auctions remain, halt script and print alert.
  if (newAuctions.length == 0) {
    ui.alert('Auction information is already up to date. Halting script.');
    return;
  }
  
  // Cache new order information.
  var newOrders = sheetNewOrders.getRange(6, 1, sheetNewOrders.getLastRow()-6, 6).getValues();
  // Clear the sheets.
  sheetLiq.getRange(3, 2, sheetLiq.getLastRow(), 13).clearContent();
  sheetOrders.getRange(2, 1, sheetOrders.getLastRow(), 12).clear();
  if (sheetOldAuctions.getLastRow()){sheetOldAuctions.getDataRange().clear()};
  // Transfer order information into manifest.
  sheetOldAuctions.getRange(1, 1, newAuctions.length, newAuctions[0].length).setValues(newAuctions);
  sheetOrders.getRange(2, 1, newOrders.length, newOrders[0].length).setValues(newOrders);
//  var AERrange = sheetOrders.getRange(2, 6, newOrders.length);
//  AERrange.copyTo(sheetOrders.getRange(2, 5, newOrders.length));
//  AERrange.clear();
  
  // Print out the number of auctions copied and the total items in those auctions.
  ui.alert('Script finished.\n\nAuctions Copied: ' + sheetOldAuctions.getLastRow()  + '\nItems In Auctions: ' + sumArray(getCol(newAuctions, 3)));
}  

function updateLiqFormat() {
  /************************************************************************
  * This script accomplishes the following tasks:
  *   1. Find order in Copy of Liq Orders Scrap from its number in Auctions
  *   2. Move order information into LIQ FORMAT with correct formatting
  *   3. Fill out all relevant formulas on the right side of LIQ FORMAT
  *   4. Adjust per item cost to align with total cost
  *   5. Repeat for each order in Auctions 
  *************************************************************************/
  
  // Set ID for the spreadsheet file to be used.
  var maniID = "1aV0bturINsUIgF-X8LcWZaWfehWEkhr6nUIVhbqdAxM";
  
  // Initialize the sheets to be accessed.
  var sheetLiq = SpreadsheetApp.openById(maniID).getSheetByName("LIQ FORMAT");
  var sheetOrders = SpreadsheetApp.openById(maniID).getSheetByName("Copy of Liq Orders Scrap");
  var sheetAuctions = SpreadsheetApp.openById(maniID).getSheetByName("Auctions");
  
  // Extract first column from Orders sheet.
  var orders = sheetOrders.getDataRange().getValues();
  var liqOrders = (getCol(orders,0));
  
  var auctions = sheetAuctions.getDataRange().getValues();
  var auctionCount = sheetAuctions.getLastRow();
  
  // Save today's properly formatted date as a variable.
  var today = new Date();
  var dd = today.getDate();
  var mm = today.getMonth()+1; // .getMonth is 0-indexed.
  var yyyy = today.getFullYear();
  if(dd<10) { dd = '0' + dd;}
  if(mm<10) { mm = '0' + mm;}
  var today = mm + '/' + dd + '/' + yyyy;
  
  // Create function that makes an array of n length and obj identical inputs.
  function rep(obj, n) {
    var arr = [[]];
    for (i=0; i < n; i++) {arr[i][0].push(obj);}
    return arr;
  }
  
  // Create function that rounds a value to exp decimal places
  function round(value, exp) {
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
  
  // Cache the order #'s already in the sheet and the first order # to be copied.
  var liq8Digit = (getCol(sheetLiq.getDataRange().getValues(),7));
  var orderID = auctions[0][0];
  
  // Initialize variables for counting/reporting purposes.
  var orderCopy = 0;
  var itemCopy = 0;
  
  // Create function that checks to see if objects from one array are contained in a second array.
  function containedIn(needles, haystack) {
    var check = [];
    for (i=0; i < needles.length; i++) {
      check[i] = haystack.indexOf(needles[i][0]) > -1;
    }
    return check.indexOf(true) > -1;
  }
         
  // If data has already been copied, output error message.
  if (containedIn(auctions, liq8Digit)) {
    SpreadsheetApp.getUi().alert('LIQ FORMAT is already up to date. Halting script.');
    return;
  }
  
  // Clear LIQ FORMAT sheet and remove empty rows.
  sheetLiq.getRange(3, 1, sheetLiq.getMaxRows()-2, sheetLiq.getLastColumn()).clear();
  sheetLiq.deleteRows(3, sheetLiq.getMaxRows()-2)
  
  for (i=0; i < auctionCount; i++) {
    // Update orderID to the next order #.
    var orderID = auctions[i][0];
    // Save item count for that order as a variable
    var itemCount = auctions[i][3];
    sheetLiq.insertRowsAfter(2, itemCount);
    // Find order in Copy of Liq Orders Scrap
    var orderIndex = liqOrders.indexOf(orderID);
    if (orderIndex == -1) {
      SpreadsheetApp.getUi().alert('Could not find order #' + orderID + '. Hit OK to continue.');
      break;
    }
    // Find last row of data.
    var liqLastRow = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("LIQ FORMAT").getLastRow();
    // Save order as range with itemCount number of rows
    var orderItems = sheetOrders.getRange(orderIndex+1, 2, itemCount, 5);
    // Copy range values over to LIQ FORMAT
    orderItems.copyValuesToRange(sheetLiq, 3, 7, liqLastRow+1, liqLastRow+itemCount+1);
    sheetLiq.getRange(liqLastRow+1, 1, itemCount, 9).setBackground('white');
      
    // Copy A/E/R from Buy Site to correct column
    var AER = sheetLiq.getRange(liqLastRow+1, 8, itemCount);
    sheetLiq.getRange(liqLastRow+1, 7, itemCount).moveTo(AER);
    var formulaRange = sheetLiq.getRange(2, 10, 1, 5);
    // Fill in date, buy site, and cost information.
    for (var j=1; j <= itemCount; j++) {
      sheetLiq.getRange(liqLastRow+j, 2).setValue(today);
      sheetLiq.getRange(liqLastRow+j, 7).setValue("LIQUIDATION");
      sheetLiq.getRange(liqLastRow+j, 9).setValue(orderID);
      formulaRange.copyTo(sheetLiq.getRange(liqLastRow+j, 10, 1, 5));
    }
    // Copy per item cost values.
    sheetLiq.getRange(liqLastRow+1, 15, itemCount).setValue(sheetLiq.getRange(liqLastRow+1, 14, itemCount).getDisplayValues());
    // Have if statement comparing rounded cost to actual cost
    var prices = sheetLiq.getRange(liqLastRow+1, 15, itemCount).getValues();
    var orderTotal = sheetLiq.getRange(liqLastRow+1,10).getValue();
    var roundedTotal = Number(round(sumArray(prices), 2));
    if (roundedTotal < orderTotal) {
      // If rounded is lower, compensate top per item cost
      sheetLiq.getRange(liqLastRow+1, 15).setValue(Number(prices[0]) + orderTotal - roundedTotal);
    }
    else if (roundedTotal > orderTotal) {
      // If rounded is higher, compensate bottom per item cost
      sheetLiq.getRange(liqLastRow+itemCount, 15).setValue(Number(prices[itemCount-1]) + orderTotal - roundedTotal);
    }
    orderCopy++;
    itemCopy = itemCopy + itemCount;
  }
  // Post dialogue box showing # of orders and items copied to LIQ FORMAT.
  SpreadsheetApp.getUi().alert('Script finished.\n\nOrders Copied: ' +  orderCopy + '\nItems Copied: ' + itemCopy);
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
  var maniID = "1aV0bturINsUIgF-X8LcWZaWfehWEkhr6nUIVhbqdAxM";
  var workID = "1okDFF9236lGc4vU6W7HOD8D-3ak8e_zntehvFatYxnI";
  var liqID = "1Xqsc6Qe_hxrWN8wRd_vgdBdrCtJXVlvVC9w53XJ0BUM";
  
  // Load the sheets between which data will be transferred.
  var sheetManifest = SpreadsheetApp.openById(maniID).getSheetByName("LIQ FORMAT");
  var sheetFuture = SpreadsheetApp.openById(workID).getSheetByName("Future Listing");
  var sheetLiquid = SpreadsheetApp.openById(liqID).getSheetByName("Liquidation Orders");
  
  // Save last row in each sheet to be used for indexing later.
  var maniLastRow = sheetManifest.getLastRow();
  var liqLastRow = sheetLiquid.getLastRow();
  
  // Load all of the values from manifest sheet.
  var maniValues = sheetManifest.getDataRange().getValues();
  
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
    // To Future(column): Title(3), ASIN(4), LPN(5), A/E/R(6), and 7-digit Order #(7) from Manifest.
    sheetFuture.getRange(i, 2).setValue(highSKU + k);       // SKU
    sheetFuture.getRange(i, 3).setValue(maniValues[i][2]);  // Title
    sheetFuture.getRange(i, 4).setValue(maniValues[i][4]);  // ASIN
    sheetFuture.getRange(i, 5).setValue(maniValues[i][5]);  // LPN
    sheetFuture.getRange(i, 6).setValue(maniValues[i][7]);  // A/E/R
    sheetFuture.getRange(i, 7).setValue(maniValues[i][9]);  // Order #
    // To Liquid(column): Date(2), Title(3), Quantity(4), ASIN(5), Buy Site(6), A/E/R(7), 7-digit #(8), Buy Price(11), and Card(12) from Manifest.
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
  }
}
