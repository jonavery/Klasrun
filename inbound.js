function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Automation Menu')
    //.addItem('Import New Blackwrap', 'importBlackwrap')
    //.addItem('Generate Price Estimates', 'generatePrices')
    //.addItem('Import Price Estimates', 'importPrices')
    //.addSeparator()
    .addItem('Update Export', 'updateExport')
    .addItem('Export to LIQ & WORK', 'exportData')
    .addSeparator()
    .addItem('Lookup ASINs', 'lookupASINs')
    .addItem('Update ASIN DB', 'updateASINs')
    .addToUi();
}

function testNono() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Research');
  var line = 6;
  var item = sheet.getRange(line, 2).getValue();
  var itemCount = 0;
  while (item != "") {
    itemCount++;
    item = sheet.getRange(line+itemCount, 2).getValue();
  }
  updateFormula(sheet, itemCount, line);
  nono(sheet, itemCount, line);
}

function designate(salePrice, weight) {
  var score = salePrice - weight;
  if (salePrice < 25 || score < 20) {
    return "R";
  }
  if (weight == "undefined" || salePrice == "undefined" || salePrice == "MANUAL") {
    return "?";
  }
  if (weight > 75) {
    if (score >= 100) {
      return "?";
    } else {
      return "R";
    }
  }
  if (score >= 100) {
    return "E";
  }
  if (weight > 30 && weight < 50) {
    if (salePrice < 100) {
      return "R";
    } else {
      return "A";
    }
  }
  if (score >= 20) {
    return "A";
  } else {
    return "R";
  }
}

function updateFormula(sheet, itemCount, line) {
  // Set summation formulas.
  sheet.getRange(itemCount+line, 1, 2, 9).setFontStyle('bold');
  sheet.getRange(itemCount+line, 1).setValue("SUBTOTAL");
  sheet.getRange(itemCount+line+1, 1).setValue("MY BUY PRICE");
  var countA1 = sheet.getRange(line, 3, itemCount).getA1Notation();
  var amazonA1 = sheet.getRange(line, 7, itemCount).getA1Notation();
  var feesA1 = sheet.getRange(line, 8, itemCount).getA1Notation();
  sheet.getRange(itemCount+line, 3).setFormula("=SUM("+countA1+")");
  sheet.getRange(itemCount+line, 7).setFormula("=SUM("+amazonA1+")");
  var feeSumA1 = sheet.getRange(itemCount+line, 8).setFormula("=SUM("+feesA1+")").getA1Notation();
  var buyA1 = sheet.getRange(itemCount+line+1, 8).setFormula("=SUM("+feeSumA1+"*0.6)").setBackgroundRGB(217, 234, 211).getA1Notation();
  sheet.getRange(itemCount+line, 9, 2).setBackgroundRGB(255, 153, 0);
  sheet.getRange(itemCount+line+1, 9).setFormula("=ROUND("+buyA1+"*0.9,2)");

  // Set vlookup formulas.
  var lastRow = sheet.getLastRow();
  var rangeA1 = sheet.getRange(itemCount+line+5, 4, lastRow-itemCount-line-4, 5).getA1Notation();
  for (var i = line; i < itemCount+line; i++) {
    var asinA1 = sheet.getRange(i, 4).getA1Notation();
    sheet.getRange(i, 6).setFormula("=VLOOKUP("+asinA1+","+rangeA1+",3,FALSE)");
    sheet.getRange(i, 7).setFormula("=VLOOKUP("+asinA1+","+rangeA1+",4,FALSE)");
    sheet.getRange(i, 8).setFormula("=VLOOKUP("+asinA1+","+rangeA1+",5,FALSE)");
    sheet.getRange(i, 10).setFormula("=VLOOKUP("+asinA1+","+rangeA1+",7,FALSE)");
    sheet.getRange(i, 11).setFormula("=VLOOKUP("+asinA1+","+rangeA1+",8,FALSE)");
    sheet.getRange(i, 12).setFormula("=VLOOKUP("+asinA1+","+rangeA1+",9,FALSE)");
  }
}

function nono(sheet, itemCount, line) {
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
  var items = sheet.getRange(line, 2, itemCount).getValues();
  for (var i = 0; i < itemCount; i++) {
    for (var j = 0; j < banBrand.length; j++) {
      if (items[i][0].indexOf(banBrand[j]) != -1) {
        sheet.getRange(line+i, 6).setValue('R');
        sheet.getRange(line+i, 1).setValue('BAN');
        Logger.log("Title: "+items[i][0]+"     Brand: "+banBrand[j]);
      }
    }
  }

  // Make items returns if they are on banned ASIN list.
  var itemASIN = sheet.getRange(line, 4, itemCount).getValues();
  for (var i = 0; i < itemCount; i++) {
    for (var j = 0; j < banASIN.length; j++) {
      if (itemASIN[i][0] == banASIN[j]) {
        sheet.getRange(line+i, line).setValue('R');
        sheet.getRange(line+i, 1).setValue('BAN');
        Logger.log("ItemASIN: "+itemASIN[i][0]+"     BanASIN: "+banASIN[j]);
      }
    }
  }
}

function importBlackwrap() {
  /**
  * This script imports a new blackwrap manifest into the research tab.
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
    if (header == "MSRP") {var msrpCol = i;}
  }
  var titles = getCol(blackwrapValues, titleCol);
  var asins = getCol(blackwrapValues, asinCol);
  var LPNs = getCol(blackwrapValues, lpnCol);
  var MSRPs = getCol(blackwrapValues, msrpCol);

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
    sheet.getRange(j, 9).setValue(MSRPs[i]);
  }

  // Set summation and VLOOKUP formulas.
  updateFormula(sheet, itemCount, 6);

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

  // Make banned items returns
//  nono(sheet, itemCount, 6);

  // Copy formula output values and paste them as text.
  var vlookupValues = sheet.getRange(6, 6, itemCount, 3).getValues();
  sheet.getRange(6, 6, itemCount, 3).setValues(vlookupValues);
}

function generatePrices() {
  /**
  * This script uses the Amazon Products API in tandem with
  * ec2-13-57-188-159.us-west-1.compute.amazonaws.com PHP scripting to create a JSON file of the
  * products in sheet with their price, weight, and sales
  * rank on Amazon.
  */

  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Enter first line of order:');

  if (response.getSelectedButton() == ui.Button.OK) {
    ui.alert(
      'Go to the following URL and wait for a success message:\n\n'
      + 'http://ec2-13-57-188-159.us-west-1.compute.amazonaws.com/AmazonMWS/MarketplaceWebServiceProducts/Functions/BlackwrapPricing.php'
      + '?line=' + response.getResponseText());
  } else {
    ui.alert(
      'Go to the following URL and wait for a success message:\n\n'
      + 'http://ec2-13-57-188-159.us-west-1.compute.amazonaws.com/AmazonMWS/MarketplaceWebServiceProducts/Functions/BlackwrapPricing.php');
  }
}

function importPrices() {
  /**
  * This script accomplishes the following tasks:
  *  1. Pull json file from MWS server
  *  2. Convert json into multidimensional array
  *  3. Update sheet with ASINs/UPCs.
  */

  var ui = SpreadsheetApp.getUi();
  var line = ui.prompt('Enter first line of order:').getResponseText();

  // Fetch the json array from website and parse into JS object.
  var response = UrlFetchApp.fetch('http://ec2-13-57-188-159.us-west-1.compute.amazonaws.com/AmazonMWS/MarketplaceWebServiceProducts/Functions/blackwrap.json');
  var json = response.getContentText();
  var data = JSON.parse(json);

  // Convert data object into multidimensional array.
  var itemCount = data.length;
  var itemArray = [];
  for (var i = 0; i < itemCount; i++) {
    var item = data[i];
    itemArray.push([
      item.Price,
      item.Rank,
      item.Weight
    ]);
  }

  // Push array into research tab.
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Research');
  sheet.getRange(line, 10, itemCount, 3).setValues(itemArray);

  // Wait for Sheets to process new values.
  var checkValue = itemArray[0][0];
  var cellValue = sheet.getRange(line, 10).getDisplayValue();
  var stopper = 0;
  while (String(checkValue) != String(cellValue) && stopper < 10000) {
    var cellValue = sheet.getRange(line, 10).getDisplayValue();
    stopper++;
  }
  // Sort items into A, E, R, and ? designations.
  researchItems(line);
}

function researchItems(line) {
  /**
  * This script utilizes the designate function to sort items into categories of:
  *   A = Amazon
  *   E = eBay
  *   R = Return (liquidate)
  *   CL = CraigsList (local listing)
  *
  * A "P" after the designation signifies permanence, and will not be edited by the code.
  */

  // Set ID for the spreadsheet file to be used.
  var inboundID = "1TaxBUL8WjTvV3DjJEMduPK6Qs3A5GoFDmZHiUcc-LUY";

  // Initialize the sheets to be accessed.
  var sheetResearch = SpreadsheetApp.openById(inboundID).getSheetByName("Research");

  // Find and cache last row of order.
  var cell = 1;
  var lastItemRow = line;
  while (cell == 1) {
    lastItemRow++;
    var cell = sheetResearch.getRange(lastItemRow+1, 3).getValue();
  }

  // Cache weights, msrp's, and current designations.
  var dataResearch = sheetResearch.getDataRange().getValues();
  var colWeight = getCol(dataResearch, 11);
  var colSalePrice = getCol(dataResearch, 9);
  var colAER = getCol(dataResearch, 5);

  // Cycle through all items and sort into A, E, R, and CL.
  for (var i = line-1; i < lastItemRow; i++) {
    var preDesig = colAER[i];
    if (preDesig.indexOf('P') > -1) {continue;}
    var postDesig = designate(colSalePrice[i], colWeight[i]);
    sheetResearch.getRange(i+1, 6).setValue(postDesig);
  }
}


function updateExport() {
  /************************************************************************
  * This script accomplishes the following tasks:
  *   1. Find order in Research from its number in Cycles
  *   2. Move order information into Export with correct formatting
  *   3. Fill out all relevant formulas on the right side of Export
  *   4. Adjust per item cost as weighted average of net profit
  *************************************************************************/

  // Prompt user for number of orders.
  var ui = SpreadsheetApp.getUi();
  var firstOrder = Number(ui.prompt('Enter number of first order row:').getResponseText()-1);
  var orderCount = Number(ui.prompt('Enter number of orders to be transferred:').getResponseText());

  // Set ID for the spreadsheet file to be used.
  var inboundID = "1TaxBUL8WjTvV3DjJEMduPK6Qs3A5GoFDmZHiUcc-LUY";

  // Initialize the sheets to be accessed.
  var sheetExp = SpreadsheetApp.openById(inboundID).getSheetByName("Export");
  var sheetResearch = SpreadsheetApp.openById(inboundID).getSheetByName("Research");
  var sheetCycles = SpreadsheetApp.openById(inboundID).getSheetByName("Cycles");

  // Sort Research sheet by Auction ID to group auctions together.
  sheetResearch.sort(1);

  // Extract first column from Research sheet and initialize order information.
  var orders = sheetResearch.getDataRange().getValues();
  var orderCol = getCol(orders,0);
  var auctions = sheetCycles.getDataRange().getValues();

  // Clear Export sheet and remove empty rows.
  sheetExp.getRange(3, 1, sheetExp.getMaxRows()-2, sheetExp.getLastColumn()).clear();
  var lastRow = sheetExp.getLastRow();
  var copyCount = 0;
  var itemCount = getCol(auctions.slice(firstOrder,firstOrder+orderCount),3);
  var itemTotal = sumArray(itemCount);
  var rowDiff = itemTotal - sheetExp.getMaxRows() + 2;
  if (rowDiff > 0) {sheetExp.insertRowsAfter(lastRow, rowDiff);}

  for (var i=0; i<orderCount; i++) {
    // Cache order ID, item count, and buy total.
    var orderID = auctions[firstOrder+i][1];
    var itemCount = auctions[firstOrder+i][3];
    var orderTotal = auctions[firstOrder+i][2];

    // Cache row positions.
    var r = lastRow + 1;

    // Find order in Research sheet
    var orderStart = orderCol.indexOf(orderID);
    if (orderStart < 0) {
      SpreadsheetApp.getUi().alert('Could not find order #:' + orderID + '. Aborting...');
      return;
    }
    var orderEnd = orderCol.lastIndexOf(orderID)+1;
    var rowCount = orderEnd-orderStart;
    var e = r+rowCount-1;
    // Save order columns as ranges with itemCount number of rows
    var orderItems = sheetResearch.getRange(orderStart+1, 2, rowCount, 4);
    var orderAERs = sheetResearch.getRange(orderStart+1, 6, rowCount);
    var orderMSRPs = sheetResearch.getRange(orderStart+1, 9, rowCount);

    // Check to see if any items need to be duplicated.
    var qtyCol = getCol(sheetResearch.getRange(orderStart+1, 3, rowCount).getValues(),0);
    var dupCheck = 0;
    var dupRows = [];
    for (var j=0; j < rowCount; j++) {
      if (Number(qtyCol[j]) > 1) {
        dupCheck++;
        dupRows.push(j);
      }
    }

    // Copy range values over to Export
    orderItems.copyValuesToRange(sheetExp, 2, 5, r, e);
    orderAERs.copyValuesToRange(sheetExp, 7, 7, r, e);
    orderMSRPs.copyValuesToRange(sheetExp, 9, 9, r, e);
    sheetExp.getRange(r, 1, rowCount, 8).setBackground('white');

    // Duplicate any items with quantity >1 and set qty to 1.
    for (var j=0; j < dupCheck; j++) {
      var itmQty = Number(qtyCol[dupRows[j]]);
      sheetExp.getRange(r+dupRows[j], 3).setValue(1);
      var rowValues = sheetExp.getRange(r+dupRows[j], 1, 1, 9).getValues();
      for (var k=1; k < itmQty; k++) {
        sheetExp.getRange(e+1, 1, 1, 9).setValues(rowValues);
        e++;
      }
    }

    // Hard-code in the formula for weighted average pricing.
    sheetExp.getRange(2, 15).setFormula("=ROUND(K2*I2/SUM(I$"+r+":I$"+e+"),2)");
    var formulaRange = sheetExp.getRange(2, 10, 1, 6);
    // Fill in date, buy site, and cost information.
    for (var j=0; j < itemCount; j++) {
      sheetExp.getRange(r+j, 1).setValue(today());
      sheetExp.getRange(r+j, 6).setValue("LIQUIDATION");
      sheetExp.getRange(r+j, 8).setValue(orderID);
      formulaRange.copyTo(sheetExp.getRange(r+j, 10, 1, 6));
    }

    // Compare rounded cost to actual cost
    var prices = sheetExp.getRange(r, 15, itemCount).getDisplayValues();
    var roundedTotal = sumArray(prices);
    if (roundedTotal != orderTotal) {
      // Compensate top per item cost
      sheetExp.getRange(r, 15).setValue(Number(prices[0]) + Number(orderTotal) - roundedTotal);
      Logger.log(Number(prices[0]) + orderTotal - roundedTotal);
    }
    var lastRow = lastRow + itemCount;
    var copyCount = copyCount + itemCount;
  }
  // Post dialogue box showing # of orders and items copied to LIQ FORMAT.
  SpreadsheetApp.getUi().alert('Script finished.\n\nItems Copied: ' + copyCount);
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
  var inboundID = "1TaxBUL8WjTvV3DjJEMduPK6Qs3A5GoFDmZHiUcc-LUY";
  var workID = "1okDFF9236lGc4vU6W7HOD8D-3ak8e_zntehvFatYxnI";
  var liqID = "1Xqsc6Qe_hxrWN8wRd_vgdBdrCtJXVlvVC9w53XJ0BUM";

  // Load the sheets between which data will be transferred.
  var sheetExport = SpreadsheetApp.openById(inboundID).getSheetByName("Export");
  var sheetFuture = SpreadsheetApp.openById(workID).getSheetByName("Future Listing");
  var sheetLiquid = SpreadsheetApp.openById(liqID).getSheetByName("Liquidation Orders");

  // Save last row in each sheet to be used for indexing later.
  var maniLastRow = sheetExport.getLastRow();
  var liqLastRow = sheetLiquid.getLastRow();

  // Load all of the values from manifest sheet.
  var inboundValues = sheetExport.getDataRange().getValues();

  // Prepare the future listings sheet for data entry.
  var futureMaxRows = sheetFuture.getMaxRows();
  sheetFuture.getRange(2, 1, futureMaxRows-1, sheetFuture.getLastColumn()).clear();
  var futureNeededRows = inboundValues.length + 1 - futureMaxRows;
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
  var liqNeededRows = inboundValues.length + liqLastRow - sheetLiquid.getMaxRows();
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
  var inboundOrderNums = [];
  for (var i=2; i < maniLastRow; i++) {
    var k = i-1;
    if (inboundOrderNums.indexOf(inboundValues[i][9]) == -1) {inboundOrderNums[i-2] = inboundValues[i][9];}
    if (allOrderNums.indexOf(inboundValues[i][9]) > -1) {
      Logger.log('Order #' + inboundValues[i][9] + ' has already been copied.');
      break;
    }
    // To Future(column): Title(3), ASIN(4), LPN(5), A/E/R(6), and 7-digit Order #(7) from liq orders.
    sheetFuture.getRange(i, 2).setValue(highSKU + k);       // SKU
    sheetFuture.getRange(i, 3).setValue(inboundValues[i][1]);  // Title
    sheetFuture.getRange(i, 4).setValue(inboundValues[i][3]);  // ASIN
    sheetFuture.getRange(i, 5).setValue(inboundValues[i][4]);  // LPN
    sheetFuture.getRange(i, 6).setValue(inboundValues[i][6]);  // A/E/R
    sheetFuture.getRange(i, 7).setValue(inboundValues[i][9]);  // Order #
    // To Liquid(column): Date(2), Title(3), Quantity(4), ASIN(5), Buy Site(6), A/E/R(7), 7-digit #(8), Buy Price(11), and Card(12) from liq orders.
    sheetLiquid.getRange(liqLastRow + k, 1).setValue(highSKU + k);       // SKU
    sheetLiquid.getRange(liqLastRow + k, 2).setValue(inboundValues[i][0]);  // Date
    sheetLiquid.getRange(liqLastRow + k, 3).setValue(inboundValues[i][1]);  // Title
    sheetLiquid.getRange(liqLastRow + k, 4).setValue(inboundValues[i][2]);  // Quantity
    sheetLiquid.getRange(liqLastRow + k, 5).setValue(inboundValues[i][3]);  // ASIN
    sheetLiquid.getRange(liqLastRow + k, 6).setValue(inboundValues[i][5]);  // Buy Site
    sheetLiquid.getRange(liqLastRow + k, 7).setValue(inboundValues[i][6]);  // A/E/R
    sheetLiquid.getRange(liqLastRow + k, 8).setValue(inboundValues[i][9]);  // Order #
    sheetLiquid.getRange(liqLastRow + k, 9).setValue("FBA");             // Sell Site
    sheetLiquid.getRange(liqLastRow + k, 10).setValue("FBA");            // Sell Order
    sheetLiquid.getRange(liqLastRow + k, 11).setValue(inboundValues[i][14]);// Buy Price
    sheetLiquid.getRange(liqLastRow + k, 12).setValue(inboundValues[i][11]);// Card
    sheetLiquid.getRange(liqLastRow + k, 18).setValue("IN HOUSE");       // Set Month to IN HOUSE
    sheetLiquid.getRange(liqLastRow + k, 29).setValue(inboundValues[i][12]);// Category
    // Setup liquidation formulas for new entry.
    var r = String(liqLastRow + k);
    sheetLiquid.getRange(liqLastRow + k, 14).setFormula("=M"+r+"-K"+r);  // Actual Profit
    sheetLiquid.getRange(liqLastRow + k, 15).setFormula("=M"+r+"/K"+r);  // Actual % Increase
    sheetLiquid.getRange(liqLastRow + k, 22).setFormula("=VLOOKUP(A"+r+",Returns!A:A,1,0)");        // RETURNS V
    sheetLiquid.getRange(liqLastRow + k, 23).setFormula("=VLOOKUP(A"+r+",Salvage!A:A,1,0)");        // SALVAGE V
    sheetLiquid.getRange(liqLastRow + k, 24).setFormula("=VLOOKUP(A"+r+",Reimbursements!F:F,1,0)"); // REIMBURSE V
    sheetLiquid.getRange(liqLastRow + k, 25).setFormula("=VLOOKUP(A"+r+",Inventory!B:B,1,0)");      // INVENTORY V
    sheetLiquid.getRange(liqLastRow + k, 26).setFormula("=VLOOKUP(A"+r+",Connor!G:H,2,0)");         // FBA SHIPMENT STATUS
    sheetLiquid.getRange(liqLastRow + k, 27).setFormula("=VLOOKUP(A"+r+",Connor!K:K,1,0)");         // FBA SHIPMENT ISSUE
  }
  highlightAER();

  // Note in cycles sheet that auction has been exported.
  var sheetCycles = SpreadsheetApp.openById(inboundID).getSheetByName("Cycles");
  var auctionCol = getCol(sheetCycles.getDataRange().getValues(),4);
  for (var i=0; i < inboundOrderNums.length; i++) {
    var auctionIndex = auctionCol.indexOf(Number(inboundOrderNums[i]));
    Logger.log(auctionIndex);
    sheetCycles.getRange(auctionIndex+1, 9).setValue(today());
  }
}

function highlightAER() {
  /**
  * This script highlights each A/E/R cell according to its designation.
  * The script is coded to leave highlighted cells/rows alone.
  */

  // Initialize sheet and save values.
  var workID = "1okDFF9236lGc4vU6W7HOD8D-3ak8e_zntehvFatYxnI";
  var sheetWork = SpreadsheetApp.openById(workID).getSheetByName("Future Listing");
  var rangeWork = sheetWork.getDataRange();
  var workValues = rangeWork.getValues();
  var workColors = rangeWork.getBackgrounds();

  // Cache A/E/R column of sheet.
  var aerValues = getCol(workValues, 5);

  // Loop through A/E/R column and color cells with a switch statement.
  for (i=1; i<aerValues.length; i++) {
    if (workColors[i][0] == "#ffffff") {
      var activeRange = sheetWork.getRange(i+1, 2, 1, 6);
      switch (aerValues[i]) {
        case 'a':
        case 'A':
        case 'AP':
          activeRange.setBackground('white');
          break;
        case 'e':
        case 'E':
        case 'EP':
          activeRange.setBackground('#ff00ff');
          break;
        case 'r':
        case 'R':
        case 'RP':
          activeRange.setBackground('orange');
          break;
        case 'RHD':
          activeRange.setBackground('blue');
          activeRange.setFontColor('white');
          break;
        case 'cl':
        case 'CL':
        case 'CL ONLY':
          activeRange.setBackground('#b43800');
          activeRange.setFontColor('white');
          break;
        default:
          activeRange.setBackground('gray');
          break;
      }
    }
  }
}

function lookupASINs() {
  /**
  * This script uses the Amazon Products API in tandem with
  * ec2-13-57-188-159.us-west-1.compute.amazonaws.com PHP scripting to create a JSON file of the
  * products in sheet with their price, weight, and sales
  * rank on Amazon.
  */

  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Enter first line of order:')

  if (response.getSelectedButton() == ui.Button.OK) {
    UrlFetchApp.fetch('http://ec2-13-57-188-159.us-west-1.compute.amazonaws.com/AmazonMWS/MarketplaceWebServiceProducts/Functions/LookupASIN.php?line='+response.getResponseText());
  } else {
    UrlFetchApp.fetch('http://ec2-13-57-188-159.us-west-1.compute.amazonaws.com/AmazonMWS/MarketplaceWebServiceProducts/Functions/LookupASIN.php');
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Research");
  var line = parseInt(response) || 6;
  var item = sheet.getRange(line, 2).getValue();
  var itemCount = 0;
  while (item != "") {
    itemCount++;
    item = sheet.getRange(line+itemCount, 2).getValue();
  }
  Logger.log(itemCount);

  updateFormula(sheet, itemCount, line);
//  nono(sheet, itemCount, line);
}

function updateASINs() {
  /************************************************************************
  * This script accomplishes the following tasks:
  *   1. Find unique items in Research
  *   2. Filter out duplicate items and unresearched items
  *   3. Move remaining items into asinDB
  *************************************************************************/

  // Set ID for the spreadsheet file to be used.
  var inboundID = "1TaxBUL8WjTvV3DjJEMduPK6Qs3A5GoFDmZHiUcc-LUY";

  // Initialize the sheets to be accessed.
  var sheetDB = SpreadsheetApp.openById(inboundID).getSheetByName("AsinDB");
  var sheetResearch = SpreadsheetApp.openById(inboundID).getSheetByName("Research");

  // Cache values from Research.
  var researchValues = sheetResearch.getDataRange().getValues();
  // Remove rows with duplicate and blank ASINs from array.
  var researchUnique = cleanArray(researchValues.slice(5),3);

  // Cache values and ASINs from database.
  var dbValues = sheetDB.getDataRange().getValues().slice(1);
  var dbASINs = getCol(dbValues, 1);
  Logger.log(researchValues);

  for (var i=0; i<researchUnique.length; i++) {
    // Compare Research items to database items.
    var res = researchUnique[i];
    // Make sure AERdesignation conforms to DB enumeration rules.
    const enum = ['A','E','R','KEEP','CL','RHD'];
    const letter = [];
    for (var j=0; j<enum.length; j++) {letter.push(enum[j][0]);}
    if (!containedIn(enum,res)) {
      for (var j=0; j<letter.length; j++) {
        if (res[5][0] == letter[j]) {res[5] = enum[j]; break;}
        else {res[5] = enum[2];}
      }
    }
    // Cache needed elements from research table.
    var row = [
        res[1],
        res[3],
        res[5],
        res[6],
        "",
        res[7],
        res[10],
        res[8]
      ];
    // Add new Research items to database array.
    if (dbASINs.indexOf(res[3]) == -1) {
      dbValues.push(row);
    // Update existing Research items in database array.
    } else {
      var index = dbASINs.indexOf(res[3]);
      dbValues[index] = row;
    }
  }
  // Move new values into database.
  var dbLastRow = sheetDB.getLastRow();
  var dbMaxRows = sheetDB.getMaxRows();
  sheetDB.getRange(2,1,dbLastRow-1,8).clear();
  var newLength = dbValues.length;
  if (dbMaxRows < newLength + 1) {
    sheetDB.insertRows(2, newLength-dbMaxRows+1);
  }
  sheetDB.getRange(2,1,newLength,8).setValues(dbValues).sort(1);
  // Update MySQL database with new database values.
  var response = UrlFetchApp.fetch('http://ec2-13-57-188-159.us-west-1.compute.amazonaws.com/AmazonMWS/MarketplaceWebServiceProducts/Functions/UpdateAsinDB.php');
}

function today() {
  // Return today's date in MM/DD/YYYY format.
  var today = new Date();
  var dd = today.getDate();
  var mm = today.getMonth()+1; // .getMonth() is 0-indexed.
  var yyyy = today.getFullYear();
  if(dd<10) { dd = '0' + dd;}
  if(mm<10) { mm = '0' + mm;}
  var today = mm + '/' + dd + '/' + yyyy;
  return today;
}

function cleanArray(dirty, key) {
  // Clean rows from a two-dimensional array that have duplicate or blank keys.
  // @param {Array} dirty - Array containing duplicate and blank values.
  // @param {int} key - 0-indexed column of the primary key used to check for duplicates.
  const found = {};
  const clean = [];
  for (var i=0; i<dirty.length; i++) {
    var item = dirty[i];
    if (item[key] && !found[item[key]]) {
      clean.push(item);
      found[item[key]] = true;
    }
  }
  return clean;
}

function getCol(matrix, col){
  // Take in a matrix and extract a column from it.
  // @param {int} Col - 0-indexed number of column to be outputted.
  var column = [];
  var l = matrix.length;
  for(var i=0; i<l; i++){
     column.push(matrix[i][col]);
  }
  return column;
}

function rep(obj, n) {
  // Make an array of n length and obj identical inputs.
  // @param {*} obj - value to be repeated inside array.
  // @param {int} n - number of times to repeat value.
  var arr = [[]];
  for (i=0; i < n; i++) {arr[i][0].push(obj);}
  return arr;
}

function sumArray(array) {
  // Find the sum of an array.
  // @param {Array} array - array of numerical values.
  var sum = 0;
  for (var i = 0; i < array.length; i++) {
    var sum = sum + Number(array[i]);
  }
  return sum;
}

function containedIn(needles, haystack) {
  // Check to see if any objects from one array are contained in a second array.
  // Outputs Boolean.
  // @param {Array} needles - array of objects to be looked for.
  // @param {Array} haystack - array to look in.
  var check = [];
  for (var i=0; i < needles.length; i++) {
    check[i] = haystack.indexOf(needles[i][0]) > -1;
  }
  return check.indexOf(true) > -1;
}

function findNeedles(needle, haystack) {
  // Find all locations where needle is contained in an array.
  // Outputs array of needle indices.
  // @param {Array} needle - object to be looked for.
  // @param {Array} haystack - array to look in.
  var needleLocs = [];
  var j = 0;
  for (var i=0; i < haystack.length; i++) {
    if (haystack[i] == needle) {
      needleLocs[j] = i;
      j++;
    }
  }
  return needleLocs;
}

function round(value, exp) {
  // Rounds a value to exp decimal places
  // @param {float} value - number to be rounded.
  // @param {int} exp - number of decimal places to round number. 0 will output an integer.
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

