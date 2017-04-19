function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Automation Menu')
    .addItem('Update Orders and Auctions', 'updateSheets')
    .addItem('Update LIQ FORMAT', 'updateLiqFormat')
    .addSeparator()
    .addItem('Transfer to LIQ & WORK', 'transferData')
    .addToUi();
}

// Create function that strips specified column from an array.
function getCol(matrix, col){
  var column = [];
  var l = matrix.length;
  for(var i=0; i<l; i++){
     column.push(matrix[i][col]);
  }
  return column;
}

function updateSheets() {
  /*
  This script accomplishes the following:
    1. Import the new auctions for Liq Orders1.ods
    2. Clear the LIQ FORMAT, Orders, and Auctions sheets from Manifest.
    3. Paste the auctions into Orders and Auctions.
  */
  
  // Set ID's for the spreadsheet files to be used.
  var maniID = "1aV0bturINsUIgF-X8LcWZaWfehWEkhr6nUIVhbqdAxM";
  var liqOrdersID = "1TaxBUL8WjTvV3DjJEMduPK6Qs3A5GoFDmZHiUcc-LUY";
  var liquidID = "1Xqsc6Qe_hxrWN8wRd_vgdBdrCtJXVlvVC9w53XJ0BUM";
  
  // Initialize the sheets to be updated.
  var sheetLiq = SpreadsheetApp.openById(maniID).getSheetByName("LIQ FORMAT");
  var sheetOrders = SpreadsheetApp.openById(maniID).getSheetByName("Copy of Liq Orders Scrap");
  var sheetAuctions = SpreadsheetApp.openById(maniID).getSheetByName("Auctions");
  
  // Initialize liq orders1.ods sheets.
  var sheetNewAuctions = SpreadsheetApp.openById(liqOrdersID).getSheetByName("AUCTION");
  var sheetNewOrders = SpreadsheetApp.openById(liqOrdersID).getSheetByName("Sheet6")
  
  // Initialize liquidation sheet and check if manifest data has been transferred.
  var sheetLiquid = SpreadsheetApp.openById(liquidID).getSheetByName("Liquidation Orders");
  var maniValues = sheetLiq.getDataRange().getValues();  
  
  // Cache all order numbers currently in liquidation sheet and check to see if data has already been transferred.
  var currentOrderNums = getCol(sheetLiquid.getRange(1, 8, sheetLiquid.getLastRow()).getValues(), 0);
  if (currentOrderNums.indexOf(maniValues[2][8]) == -1) {
    SpreadsheetApp.getUi().alert('LIQ FORMAT has not been transferred to LIQ and WORK yet. Transfer data before updating auctions.');
  
  } else {
    // Check to see if auctions are up to date.
    var oldOrderNum = sheetAuctions.getRange("A1").getValue();
    var newOrderNum = sheetNewAuctions.getRange("K3").getValue();
    if (oldOrderNum != newOrderNum) {
      // Find old order in liq orders1.ods sheet and cache index.
      if (oldOrderNum > 0) {
        var orders = sheetNewOrders.getDataRange().getValues();
        var liqOrders = (getCol(orders,0));
        var orderIndex = liqOrders.indexOf(oldOrderNum);
        // Cache new order information.
        var newOrders = sheetNewOrders.getRange(6, 1, orderIndex-6, 5).getValues();
      } else {
        var newOrders = sheetNewOrders.getRange(6, 1, sheetNewOrders.getLastRow()-6, 5).getValues();
      }
      var newAuctions = sheetNewAuctions.getRange(3, 11, 30, 6).getValues();
      
      // Clear the sheets.
      sheetLiq.getRange(3, 2, sheetLiq.getLastRow(), 13).clearContent();
      sheetOrders.getRange(2, 1, sheetOrders.getLastRow(), 12).clear();
      if (oldOrderNum>0) {sheetAuctions.getRange(1, 1, sheetAuctions.getLastRow(), 7).clear();}
      
      // Transfer order information into manifest.
      sheetAuctions.getRange(1, 1, newAuctions.length, newAuctions[0].length).setValues(newAuctions);
      sheetOrders.getRange(2, 1, newOrders.length, newOrders[0].length).setValues(newOrders);
      
    // Stop script if auction information is already up to date.
    } else if (oldOrderNum == newOrderNum) {
      SpreadsheetApp.getUi().alert('Auction information is already up to date. Halting script.');
    }
  }
}


function updateLiqFormat() {
  /*
  
  This script accomplishes the following tasks:
  1. Find order in Copy of Liq Orders Scrap from its number in Auctions
  2. Move order information into LIQ FORMAT with correct formatting
  3. Fill out all relevant formulas on the right side of LIQ FORMAT
  4. Adjust per item cost to align with total cost
  5. Repeat for each order in Auctions
  
  */
  
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
  var auctionCount = sheetAuctions.getDataRange().getNumRows();
  
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
  
  // Create function that finds the sum of an array.
  function sumArray(array) {
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
  } else {
    for (i=0; i < auctionCount; i++) {
      // Update orderID to the next order #.
      var orderID = auctions[i][0];
      // Save item count for that order as a variable
      var itemCount = auctions[i][3];
      // Find order in Copy of Liq Orders Scrap
      var orderIndex = liqOrders.indexOf(orderID);
      if (orderIndex == -1) {
        SpreadsheetApp.getUi().alert('Could not find order #' + orderID + '. Hit OK to continue.');
      } else {
        // Find last row of data.
        var liqLastRow = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("LIQ FORMAT").getLastRow();
        // Save order as range with itemCount number of rows
        var orderItems = sheetOrders.getRange(orderIndex+1, 2, itemCount, 4);
        // Copy range values over to LIQ FORMAT
        orderItems.copyValuesToRange(sheetLiq, 3, 6, liqLastRow+1, liqLastRow+itemCount+1)
        
        // Copy A/E/R from Buy Site to correct column
        var AER = sheetLiq.getRange(liqLastRow+1, 7, itemCount);
        sheetLiq.getRange(liqLastRow+1, 6, itemCount).moveTo(AER);
        var formulaRange = sheetLiq.getRange(2, 9, 1, 5);
        // Fill in date and buy site information.
        for (j=0; j < itemCount; j++) {
          sheetLiq.getRange(liqLastRow+1+j, 2).setValue(today);
          sheetLiq.getRange(liqLastRow+1+j, 6).setValue("LIQUIDATION");
          sheetLiq.getRange(liqLastRow+1+j, 8).setValue(orderID);
          formulaRange.copyTo(sheetLiq.getRange(liqLastRow+1+j, 9, 1, 5));
        }
        
        // Copy per item cost values.
        sheetLiq.getRange(liqLastRow+1, 14, itemCount).setValue(sheetLiq.getRange(liqLastRow+1, 13, itemCount).getDisplayValues());
        // Have if statement comparing rounded cost to actual cost
        var prices = sheetLiq.getRange(liqLastRow+1, 14, itemCount).getValues()
        var orderTotal = sheetLiq.getRange(liqLastRow+1,10).getValue()
        var roundedTotal = Number(round(sumArray(prices), 2));
        if (roundedTotal < orderTotal) {
          // If rounded is lower, compensate top per item cost
          sheetLiq.getRange(liqLastRow+1, 14).setValue(Number(prices[0]) + orderTotal - roundedTotal);
        }
        else if (roundedTotal > orderTotal) {
          // If rounded is higher, compensate bottom per item cost
          sheetLiq.getRange(liqLastRow+itemCount, 14).setValue(Number(prices[itemCount-1]) + orderTotal - roundedTotal);
        }
      }
    }
  }
}

function transferData() {
  /*
  This function will accomplish the following:
    1. Load the relevant Liquidation, Work, and Manifest sheets.
    2. Save the needed ranges from the Manifest sheet.
    3. Copy over the ranges into the Liquidation and Work sheets with the proper formatting.
    4. Fill in any constant values in the Liquidation and Work Sheets.
  */
  
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
  var futureLastRow = sheetFuture.getLastRow();
  var liqLastRow = sheetLiquid.getLastRow();
  
  // Load all of the values from manifest sheet.
  var maniValues = sheetManifest.getDataRange().getValues();  
  
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
  if (allOrderNums.indexOf(maniValues[2][8]) == -1) {
    for (i=2; i < maniLastRow-2; i++) {   
      var k = i-1;
      // To Future(column): Title(3), UPC(4), A/E/R(5), and 7-digit Order #(6) from Manifest.
      sheetFuture.getRange(futureLastRow + k, 2).setValue(highSKU + k);
      sheetFuture.getRange(futureLastRow + k, 3).setValue(maniValues[i][2]);
      sheetFuture.getRange(futureLastRow + k, 4).setValue(maniValues[i][4]);
      sheetFuture.getRange(futureLastRow + k, 5).setValue(maniValues[i][6]);
      sheetFuture.getRange(futureLastRow + k, 6).setValue(maniValues[i][8]);
      // To Liquid(column): Date(2), Title(3), Quantity(4), UPC(5), Buy Site(6), A/E/R(7), 7-digit #(8), Buy Price(11), and Card(12) from Manifest.
      sheetLiquid.getRange(liqLastRow + k, 1).setValue(highSKU + k);
      sheetLiquid.getRange(liqLastRow + k, 2).setValue(maniValues[i][1]);
      sheetLiquid.getRange(liqLastRow + k, 3).setValue(maniValues[i][2]);
      sheetLiquid.getRange(liqLastRow + k, 4).setValue(maniValues[i][3]);
      sheetLiquid.getRange(liqLastRow + k, 5).setValue(maniValues[i][4]);
      sheetLiquid.getRange(liqLastRow + k, 6).setValue(maniValues[i][5]);
      sheetLiquid.getRange(liqLastRow + k, 7).setValue(maniValues[i][6]);
      sheetLiquid.getRange(liqLastRow + k, 8).setValue(maniValues[i][8]);
      sheetLiquid.getRange(liqLastRow + k, 9).setValue("FBA");
      sheetLiquid.getRange(liqLastRow + k, 10).setValue("FBA");
      sheetLiquid.getRange(liqLastRow + k, 11).setValue(maniValues[i][13]);
      sheetLiquid.getRange(liqLastRow + k, 12).setValue(maniValues[i][10]);
     }
  } else {
     SpreadsheetApp.getUi().alert('The data has already been copied. Halting script to avoid duplicity.');
  }
}
