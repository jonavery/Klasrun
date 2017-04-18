function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Automation Menu')
    .addItem('Auctions >>> LIQ FORMAT', 'automate')
    .addItem('Transfer to LIQ & WORK', 'transferData')
    .addToUi()
}

function automate() {
  /*
  
  This script accomplishes the following tasks:
  1. Find order in Copy of Liq Orders Scrap from its number in Auctions
  2. Move order information into LIQ FORMAT with correct formatting
  3. Fill out all relevant formulas on the right side of LIQ FORMAT
  4. Adjust per item cost to align with total cost
  5. Repeat for each order in Auctions
  
  */
  
  var sheetOrders = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Copy of Liq Orders Scrap");
  var sheetAuctions = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Auctions");
  var sheetLiq = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("LIQ FORMAT");
  
  // Create function that strips specified column from an array.
  function getCol(matrix, col){
    var column = [];
    var l = matrix.length;
    for(var i=0; i<l; i++){
       column.push(matrix[i][col]);
    }
    return column;
  }
  
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
  
  // Check to see if the first order has already been copied.
  if (liq8Digit.indexOf(orderID) == -1) {
    for (i=0; i < auctionCount; i++) {
      // Find last row of data.
      var liqLastRow = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("LIQ FORMAT").getLastRow();
      
      var orderID = auctions[i][0];
      // Save item count for that order as a variable
      var itemCount = auctions[i][5];
      // Find order in Copy of Liq Orders Scrap
      var orderIndex = liqOrders.indexOf(orderID);
      Logger.log(orderIndex);
      Logger.log(itemCount);
      
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
      Logger.log("Order Total: " + orderTotal);
      Logger.log("Pre Item Sum: " + roundedTotal);
      if (roundedTotal < orderTotal) {
        // If rounded is lower, compensate top per item cost
        sheetLiq.getRange(liqLastRow+1, 14).setValue(Number(prices[0]) + orderTotal - roundedTotal);
        }
      else if (roundedTotal > orderTotal) {
        // If rounded is higher, compensate bottom per item cost
        sheetLiq.getRange(liqLastRow+itemCount, 14).setValue(Number(prices[itemCount-1]) + orderTotal - roundedTotal);
        }  
    }
  // If data has already been copied, output error message.
  } else {
    SpreadsheetApp.getUi().alert('The data has already been copied. Halting script to avoid duplicity.');
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
  
  // Manifest (TEST) ID: 1aoaePjfMCQ8eRwWlLkzhekq5Bom-vg9Cr7sWoqNX6PA
  // Work (TEST) ID: 1w28MV69JaR99e2m-2hveMLDY_Ukbz1Meg-9RQK0-sik
  // Liquidation (TEST) ID: 1Vfzm-MogYyJB88YFmA4oWQ1M4bpT4Onw-ZGeFSpqseo
  
  
  // Load the sheets between which data will be transferred.
  var sheetManifest = SpreadsheetApp.openById("1aoaePjfMCQ8eRwWlLkzhekq5Bom-vg9Cr7sWoqNX6PA").getSheetByName("LIQ FORMAT");
  var sheetFuture = SpreadsheetApp.openById("1w28MV69JaR99e2m-2hveMLDY_Ukbz1Meg-9RQK0-sik").getSheetByName("Future Listing");
  var sheetLiqOrders = SpreadsheetApp.openById("1Vfzm-MogYyJB88YFmA4oWQ1M4bpT4Onw-ZGeFSpqseo").getSheetByName("Liquidation Orders");
  
  // Save last row in each sheet to be used for indexing later.
  var maniLastRow = sheetManifest.getLastRow();
  var futureLastRow = sheetFuture.getLastRow();
  var liqLastRow = sheetLiqOrders.getLastRow();
  
  // Load all of the values from manifest sheet.
  var maniValues = sheetManifest.getDataRange().getValues();  
  
  // Load last SKU # from liquidation sheet. ***POTENTIAL PROBLEM IF LAST ROW'S SKU IS NOT HIGHEST SKU***
  var lastSKU = sheetLiqOrders.getRange(liqLastRow, 1).getValue();
  
  // Create function that strips specified column from an array.
  function getCol(matrix, col){
    var column = [];
    var l = matrix.length;
    for(var i=0; i<l; i++){
       column.push(matrix[i][col]);
    }
    return column;
  }
  
  var allOrderNums = getCol(sheetLiqOrders.getRange(1, 8, liqLastRow).getValues(), 0);
  Logger.log(allOrderNums);
  Logger.log(maniValues[2][8]);
  Logger.log(allOrderNums.indexOf(maniValues[2][8]) == -1);
  
  if (allOrderNums.indexOf(maniValues[2][8]) == -1) {
    for (i=2; i < maniLastRow-2; i++) {   
      var k = i-1;
      // To Future(column): Title(3), UPC(4), A/E/R(5), and 7-digit Order #(6) from Manifest.
      sheetFuture.getRange(futureLastRow + k, 2).setValue(lastSKU + k);
      sheetFuture.getRange(futureLastRow + k, 3).setValue(maniValues[i][2]);
      sheetFuture.getRange(futureLastRow + k, 4).setValue(maniValues[i][4]);
      sheetFuture.getRange(futureLastRow + k, 5).setValue(maniValues[i][6]);
      sheetFuture.getRange(futureLastRow + k, 6).setValue(maniValues[i][8]);
      // To Liquid(column): Date(2), Title(3), Quantity(4), UPC(5), Buy Site(6), A/E/R(7), 7-digit #(8), Buy Price(11), and Card(12) from Manifest.
      sheetLiqOrders.getRange(liqLastRow + k, 1).setValue(lastSKU + k);
      sheetLiqOrders.getRange(liqLastRow + k, 2).setValue(maniValues[i][1]);
      sheetLiqOrders.getRange(liqLastRow + k, 3).setValue(maniValues[i][2]);
      sheetLiqOrders.getRange(liqLastRow + k, 4).setValue(maniValues[i][3]);
      sheetLiqOrders.getRange(liqLastRow + k, 5).setValue(maniValues[i][4]);
      sheetLiqOrders.getRange(liqLastRow + k, 6).setValue(maniValues[i][5]);
      sheetLiqOrders.getRange(liqLastRow + k, 7).setValue(maniValues[i][6]);
      sheetLiqOrders.getRange(liqLastRow + k, 8).setValue(maniValues[i][8]);
      sheetLiqOrders.getRange(liqLastRow + k, 9).setValue("FBA");
      sheetLiqOrders.getRange(liqLastRow + k, 10).setValue("FBA");
      sheetLiqOrders.getRange(liqLastRow + k, 11).setValue(maniValues[i][13]);
      sheetLiqOrders.getRange(liqLastRow + k, 12).setValue(maniValues[i][10]);
     }
  } else {
     SpreadsheetApp.getUi().alert('The data has already been copied. Halting script to avoid duplicity.');
  }
}
