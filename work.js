function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Automation Menu')
    .addItem('Get Next SKU From Liquidation', 'newSKU')
    .addItem('Update All Work Items in Liquidation', 'bulkUpdateLiquid')
    .addSeparator()
    .addItem('Transfer Returns and Salvages to Liquidation', 'liqTransfer')
//    .addSubMenu(ui.createMenu('Generate MWS item array')
//      // .addItem('Standard Small Parcel', 'createMWS')
//      .addItem('Palleted', 'palletMWS')
//      .addItem('Electronics', 'electronicsMWS'))
//    .addItem('Populate MWS Tab', 'populateMWS')
//    .addItem('Post Listings', 'postListings')
//    .addSeparator()
//    .addSubMenu(ui.createMenu('Create Shipments')
//      // .addItem('Small Parcel', 'createShipments')
//      .addItem('LTL (Palleted)', 'shipLTL')
//      .addItem('Electronics', 'shipElectronics'))
    .addToUi()
}

function todayDate() {
// Return today's properly formatted date.
  var today = new Date();
  var dd = today.getDate();
  var mm = today.getMonth()+1; // .getMonth is 0-indexed.
  var yyyy = today.getFullYear();
  if(dd<10) { dd = '0' + dd;}
  if(mm<10) { mm = '0' + mm;}
  var today = mm + '/' + dd + '/' + yyyy;
  return today;
}

function makeArray(w, h, val) {
// Create array with 'w' columns, 'h' rows, and filled with 'val'
  var arr = [];
  for(i = 0; i < h; i++) {
    arr[i] = [];
    for(j = 0; j < w; j++) {
      arr[i][j] = val;
    }
  }
  return arr;
}

function getCol(matrix, col){
// Take in a matrix and slice off a column from it.
// param Col is 0-indexed.
  var column = [];
  var l = matrix.length;
  for(var i=0; i<l; i++){
     column.push(matrix[i][col]);
  }
  return column;
}

function dailyGoal() {
  /**
  * This script updates the daily end-of-day goal in the Work sheet.
  * It runs automatically once every day at 5am.
  */

  // Cache spreadsheet ID's.
  var inboundID = "1TaxBUL8WjTvV3DjJEMduPK6Qs3A5GoFDmZHiUcc-LUY";
  var workID = "1okDFF9236lGc4vU6W7HOD8D-3ak8e_zntehvFatYxnI";

  // Initialize Work and Liquidation sheets.
  var sheetListings = SpreadsheetApp.openById(workID).getSheetByName('Listings');
  var sheetGoals = SpreadsheetApp.openById(inboundID).getSheetByName('DailyGoals');

  // Format dates in Goals sheet.
  var goalValues = sheetGoals.getDataRange().getDisplayValues();
  var goalDates = getCol(goalValues, 0);
  var goalDatesFormatted = [String(goalDates[0])];
  var months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  for (var i=1; i < goalDates.length; i++) {
    goalDates[i] = String(goalDates[i]);
    var month = months.indexOf(goalDates[i].substr(4,3))+1;
    if(month<10) {month = '0' + month;}
    var day = goalDates[i].substr(8,2);
    var year = goalDates[i].substr(11,4);
    var dateFormatted = month + '/' + day + '/' + year;
    goalDatesFormatted.push(dateFormatted);
  }

  // Cache today's date and find it in the Daily Goals sheet.
  var date = todayDate();
  var goalLastRow = sheetGoals.getLastRow();
  var goalIndex = goalDatesFormatted.indexOf(date);
  if (goalIndex == -1) {
    Logger.log(date + " not found in Daily Goals sheet.");
    return;
  }

  // Update yesterday's remaining items.
  var remaining = sheetListings.getRange(1, 1).getValue();
  sheetGoals.getRange(goalIndex, 4).setValue(remaining);


  // Cache today's goal and update it in the Work sheet.
  var goals = getCol(goalValues, 3);
  sheetListings.getRange(2, 1).setValue(goals[goalIndex]);
}

function liqTransfer() {
  /**
  * This script transfers all of the returns and salvages into the liquidation
  * and Archive sheets.
  */

  // Cache spreadsheet ID's
  var liquidID = "1Xqsc6Qe_hxrWN8wRd_vgdBdrCtJXVlvVC9w53XJ0BUM";
  var workID = "1okDFF9236lGc4vU6W7HOD8D-3ak8e_zntehvFatYxnI";
  var inboundID = "1TaxBUL8WjTvV3DjJEMduPK6Qs3A5GoFDmZHiUcc-LUY";

  // Initialize the sheets to be used.
  var sheetListings = SpreadsheetApp.openById(workID).getSheetByName('Listings');
  var sheetArchive = SpreadsheetApp.openById(workID).getSheetByName('Archive');
  var sheetReturns = SpreadsheetApp.openById(liquidID).getSheetByName('Returns');
  var sheetSalvage = SpreadsheetApp.openById(liquidID).getSheetByName('Salvage');
  var sheetResearch = SpreadsheetApp.openById(inboundID).getSheetByName('Research');

  // Cache SKU's, return groups, and salvages from listings sheet.
  var valuesListings = sheetListings.getDataRange().getValues();
  var lastRowListings = sheetListings.getLastRow();
  var lastRowArchive = sheetArchive.getLastRow();
  var SKUs = getCol(valuesListings, 1);
  var returnGroups = getCol(valuesListings, 9);
  var salvages = getCol(valuesListings, 10);

  // Cache SKU's from liquidation returns and salvages sheets.
  var valuesReturns = sheetReturns.getDataRange().getValues();
  var returnSKUs = getCol(valuesReturns, 0);
  var lastRowReturns = sheetReturns.getLastRow();
  var valuesSalvage = sheetSalvage.getDataRange().getValues();
  var salvageSKUs = getCol(valuesSalvage, 0);
  var lastRowSalvage = sheetSalvage.getLastRow();

  // Cache ASIN's and MSRP's from research sheet.
  var valuesResearch = sheetResearch.getDataRange().getValues();
  var researchASINs = getCol(valuesResearch, 3);
  var researchMSRPs = getCol(valuesResearch, 8);

  // Setup month to be input into sheets.
  var months = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"];
  var today = new Date();
  var mmm = months[today.getMonth()];
  var yyyy = String(today.getFullYear());
  var yy = yyyy.substr(-2);
  var month = mmm + ". '" + yy;

  // Iterate through each return and salvage SKU in Listings and move to appropriate sheet.
  for (var i=5; i < lastRowListings; i++) {
    if (returnGroups[i] == "" && salvages[i] == "") {continue;}
    var SKU = SKUs[i];
    Logger.log(SKU+" ");
    // Move Listing row to Archive.
    sheetArchive.insertRowAfter(lastRowArchive);
    lastRowArchive++;
//    sheetListings.getRange(i+1, 1, 1, 18).copyFormatToRange(sheetArchive, 1, 18, lastRowArchive, lastRowArchive);
//    sheetListings.getRange(i+1, 1, 1, 18).copyValuesToRange(sheetArchive, 1, 18, lastRowArchive, lastRowArchive);
//    sheetListings.deleteRow(i+1);

    if (returnSKUs.indexOf(SKU) > -1 || salvageSKUs.indexOf(SKU) > -1) {continue;}
//    Logger.log(SKU+"\n");
    if (returnGroups[i] == "") {
      // Move SKU to new row in salvage sheet.
      sheetSalvage.insertRowAfter(lastRowSalvage);
      lastRowSalvage++;
      var r = lastRowSalvage;
      sheetSalvage.getRange(r, 1).setValue(SKU);

      // Set formulae
      sheetSalvage.getRange(r, 2).setFormula("=VLOOKUP(A"+r+",'Liquidation Orders'!A:C,3,0)");
      sheetSalvage.getRange(r, 3).setFormula("=COUNTIF(A"+r+",A:A)");
      sheetSalvage.getRange(r, 4).setFormula("=VLOOKUP(A"+r+",'Liquidation Orders'!A:K,11,0)");
      sheetSalvage.getRange(r, 5).setValue(month);
    }

    if (salvages[i] == "") {
      // Move SKU to new row in return sheet.
      sheetReturns.insertRowAfter(lastRowReturns);
      lastRowReturns++;
      var r = lastRowReturns;
      var rUP = r-1;
      sheetReturns.getRange(r, 1).setValue(SKU);

      // Set formulae
      var rangeMSRP = sheetReturns.getRange(r,4);
      var rangeASIN = sheetReturns.getRange(r,5);
      sheetReturns.getRange(r, 2).setFormula("=VLOOKUP(A"+r+",'Liquidation Orders'!A:C,3,0)");
      sheetReturns.getRange(r, 3).setValue(1);
      rangeMSRP.setFormula("=INDEX($D$2:$D"+rUP+",MATCH(E"+r+",$E$2:$E"+rUP+",0))");
      rangeASIN.setFormula("=VLOOKUP(A"+r+",'Liquidation Orders'!A:E,5,0)");
      sheetReturns.getRange(r, 6).setFormula("=COUNTIF(A:A,A"+r+")");
      sheetReturns.getRange(r, 7).setFormula("=VLOOKUP(A"+r+",'Liquidation Orders'!A:L,11,0)");
      sheetReturns.getRange(r, 8).setValue(month);
      sheetReturns.getRange(r, 9).setValue(returnGroups[i]);

      // Lookup MSRP if not found
//      var ASIN = rangeASIN.getDisplayValue();
//      if (rangeMSRP.getDisplayValue() == "#N/A" && ASIN != "#N/A") {
//        var indexResearch = researchASINs.indexOf(ASIN);
//        var MSRP = researchMSRPs[indexResearch];
//        if (MSRP == "") {break;}
//        rangeMSRP.setValue(MSRP);
//      }

      // CopyPaste MSRP to make literal value instead of formula.
//      rangeMSRP.setValue(rangeMSRP.getDisplayValue());

    }
  }
  // Sort Returns sheet by Cutoff Group (Column I/9).
  sheetReturns.sort(9);

}

function newSKU() {
  /**
  * This script posts the next new SKU number by checking the Liquidation sheet
  * for the current maximum SKU and incrementing it by one.
  */

  // Initialize sheets.
  var sheetLiquid = SpreadsheetApp.openById("1Xqsc6Qe_hxrWN8wRd_vgdBdrCtJXVlvVC9w53XJ0BUM").getSheetByName("Liquidation Orders");
  var liqLastRow = sheetLiquid.getLastRow();

  // Load highest SKU # from liquidation sheet.
  var allSKUs = getCol(sheetLiquid.getRange(2, 1, liqLastRow-1).getValues(), 0);
  var highSKU = 1;
  for (i=0; i < allSKUs.length; i++) {
    if (allSKUs[i] > highSKU) {
      var highSKU = allSKUs[i];
    }
  }

  // Show SKU to user.
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Use this SKU for new item: ' + String(highSKU + 1));
}

function updateBySKU() {
  /**
  * This script gets an SKU from the user and updates the item in
  * Liquidation to match it.
  */

  // Cache spreadsheet ID's
  var liquidID = "1Xqsc6Qe_hxrWN8wRd_vgdBdrCtJXVlvVC9w53XJ0BUM";
  var workID = "1okDFF9236lGc4vU6W7HOD8D-3ak8e_zntehvFatYxnI";

  // Initialize Work and Liquidation sheets.
  var sheetListings = SpreadsheetApp.openById(workID).getSheetByName('Listings');
  var sheetLiquid = SpreadsheetApp.openById(liquidID).getSheetByName("Liquidation Orders");

  // Prompt user for SKU.
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Enter the SKU:');

  if (response.getSelectedButton() == ui.Button.OK) {
    // Cache user-given SKU and initialize work and liquidation SKU's.
    var sku = parseInt(response.getResponseText());
    var workLastRow = sheetListings.getLastRow();
    var workSKU = getCol(sheetListings.getRange(1, 2, workLastRow).getValues(), 0);
    var workValues = sheetListings.getDataRange().getValues();
    var liqLastRow = sheetLiquid.getLastRow();
    var liquidSKU = getCol(sheetLiquid.getRange(1, 1, liqLastRow).getValues(), 0);
    var liquidUPC = getCol(sheetLiquid.getRange(1, 5, liqLastRow).getValues(), 0);

   // Find index of SKU in work and liquidation.
   var workIndex = workSKU.indexOf(sku);
   var i = workIndex;
   var liquidIndex = liquidSKU.indexOf(sku);
   var title = sheetLiquid.getRange(liquidIndex+1, 3).getValue();

   // Check if title is blank or already up to date.
   if (workValues[i][2] == "" || workValues[i][3] == liquidUPC[liquidIndex]) {return;}

   if (liquidIndex == -1) {
     liquidIndex = liqLastRow;
     var r = String(liquidIndex + 1);
     sheetLiquid.insertRowAfter(liqLastRow);

     // Enter values from Work sheet.
     sheetLiquid.getRange(r, 1).setValue(workValues[i][1]); // SKU
     sheetLiquid.getRange(r, 2).setValue(todayDate());      // Date
     sheetLiquid.getRange(r, 4).setValue("1");              // Quantity
     sheetLiquid.getRange(r, 6).setValue("LIQUIDATION");    // Buy Site
     sheetLiquid.getRange(r, 8).setValue(workValues[i][6]); // Buy Order

     // Enter FBA information for new entry.
     sheetLiquid.getRange(r, 9).setValue("FBA");             // Sell Site
     sheetLiquid.getRange(r, 10).setValue("FBA");            // Sell Order
     sheetLiquid.getRange(r, 11).setValue("0.01");           // Buy Price

     // Setup liquidation formulas for new entry.
     sheetLiquid.getRange(r, 14).setFormula("=M"+r+"-K"+r);  // Actual Profit
     sheetLiquid.getRange(r, 15).setFormula("=M"+r+"/K"+r);  // Actual % Increase
     sheetLiquid.getRange(r, 22).setFormula("=VLOOKUP(A"+r+",Returns!A:A,1,0)");        // RETURNS V
     sheetLiquid.getRange(r, 23).setFormula("=VLOOKUP(A"+r+",Salvage!A:A,1,0)");        // SALVAGE V
     sheetLiquid.getRange(r, 24).setFormula("=VLOOKUP(A"+r+",Reimbursements!F:F,1,0)"); // REIMBURSE V
     sheetLiquid.getRange(r, 25).setFormula("=VLOOKUP(A"+r+",Inventory!B:B,1,0)");      // INVENTORY V
//     sheetLiquid.getRange(liqLastRow + k, 26).setFormula("=VLOOKUP(A"+r+",Connor!G:H,2,0)");         // FBA SHIPMENT STATUS
//     sheetLiquid.getRange(liqLastRow + k, 27).setFormula("=VLOOKUP(A"+r+",Connor!K:K,1,0)");         // FBA SHIPMENT ISSUE
   }

   // Copy work information into liquidation.
   sheetLiquid.getRange(liquidIndex+1, 3).setValue(workValues[i][2]);
   sheetLiquid.getRange(liquidIndex+1, 5).setValue(workValues[i][3]);
   sheetLiquid.getRange(liquidIndex+1, 7).setValue(workValues[i][5]);

   // Show changes to user.
   ui.alert(
   'Item title updated from "' + title + '" to "' + workValues[i][2] + '".');
  }
}

function bulkUpdateLiquid() {
  /**
  * This script synchronizes values in the Liquidation sheet
  * to match those in the Work sheet.
  *
  * It runs automatically once every hour.
  */

  // Cache spreadsheet ID's
  var liquidID = "1Xqsc6Qe_hxrWN8wRd_vgdBdrCtJXVlvVC9w53XJ0BUM";
  var workID = "1okDFF9236lGc4vU6W7HOD8D-3ak8e_zntehvFatYxnI";

  // Initialize Work and Liquidation sheets.
  var sheetListings = SpreadsheetApp.openById(workID).getSheetByName('Listings');
  var sheetLiquid = SpreadsheetApp.openById(liquidID).getSheetByName("Liquidation Orders");

  // Cache user-given SKU and initialize work and liquidation SKU's.
  var workLastRow = sheetListings.getLastRow();
  var workSKU = getCol(sheetListings.getRange(1, 2, workLastRow).getValues(), 0);
  var workValues = sheetListings.getDataRange().getValues();
  var liqLastRow = sheetLiquid.getLastRow();
  var liquidSKU = getCol(sheetLiquid.getRange(1, 1, liqLastRow).getValues(), 0);
  var liquidUPC = getCol(sheetLiquid.getRange(1, 5, liqLastRow).getValues(), 0);

  // Initialize counting variables.
  var updated = 0;
  var notUpdated = 0;
  var created = 0;

  for (var i = 4; i < workSKU.length; i++) {
    // Find index of SKU in work and liquidation.
    var sku = Number(workSKU[i]);
    if (isNaN(sku) || sku == "") {continue;}
    var workIndex = workSKU.indexOf(sku);
    var liquidIndex = liquidSKU.indexOf(sku);

    // Check if title is blank or already up to date.
    if (workValues[i][2] == "" || workValues[i][3] == liquidUPC[liquidIndex]) {
      notUpdated++;
      continue;
    }

    if (liquidIndex == -1) {
      created++;
      liquidIndex = liqLastRow;
      var r = String(liquidIndex + 1);
      sheetLiquid.insertRowAfter(liqLastRow);
      sheetLiquid.getRange(r, 1, 1, 27).clearFormat();

      // Enter values from Work sheet.
      sheetLiquid.getRange(r, 1).setValue(workValues[i][1]);
      sheetLiquid.getRange(r, 2).setValue(todayDate());
      sheetLiquid.getRange(r, 4).setValue("1");
      sheetLiquid.getRange(r, 6).setValue("LIQUIDATION");
      sheetLiquid.getRange(r, 8).setValue(workValues[i][6]);

      // Enter FBA information for new entry.
      sheetLiquid.getRange(r, 9).setValue("FBA");             // Sell Site
      sheetLiquid.getRange(r, 10).setValue("FBA");            // Sell Order
      sheetLiquid.getRange(r, 11).setValue("0.01");           // Buy Price

      // Setup liquidation formulas for new entry.
      sheetLiquid.getRange(r, 14).setFormula("=M"+r+"-K"+r);  // Actual Profit
      sheetLiquid.getRange(r, 15).setFormula("=M"+r+"/K"+r);  // Actual % Increase
      sheetLiquid.getRange(r, 22).setFormula("=VLOOKUP(A"+r+",Returns!A:A,1,0)");        // RETURNS V
      sheetLiquid.getRange(r, 23).setFormula("=VLOOKUP(A"+r+",Salvage!A:A,1,0)");        // SALVAGE V
      sheetLiquid.getRange(r, 24).setFormula("=VLOOKUP(A"+r+",Reimbursements!F:F,1,0)"); // REIMBURSE V
      sheetLiquid.getRange(r, 25).setFormula("=VLOOKUP(A"+r+",Inventory!B:B,1,0)");      // INVENTORY V
//       sheetLiquid.getRange(liqLastRow + k, 26).setFormula("=VLOOKUP(A"+r+",Connor!G:H,2,0)");         // FBA SHIPMENT STATUS
//       sheetLiquid.getRange(liqLastRow + k, 27).setFormula("=VLOOKUP(A"+r+",Connor!K:K,1,0)");         // FBA SHIPMENT ISSUE
   }

    // Copy work information into liquidation.
    sheetLiquid.getRange(liquidIndex+1, 3).setValue(workValues[i][2]);
    sheetLiquid.getRange(liquidIndex+1, 5).setValue(workValues[i][3]);
    sheetLiquid.getRange(liquidIndex+1, 7).setValue(workValues[i][5]);
    // Check if SKU is in liquidation sheet.
    if (liquidIndex == liqLastRow) {liqLastRow++; continue;}
      updated++;
    }
  Logger.log(
    'Items updated: ' + updated
    + '\nItems already up to date: ' + notUpdated
    + '\nItems created: ' + created);
}

function createMWS() {
  /**
  * This script uses the Amazon Products API in tandem with
  * http://ec2-13-57-188-159.us-west-1.compute.amazonaws.com PHP scripting to create a JSON file of the
  * items in the SCRAP sheet currently waiting to be listed.
  *
  * Use this function in tandem with the populateMWS() and
  * postListings() functions to list products on Amazon.
  */

  SpreadsheetApp.getUi().alert(
    'Go to the following URL and wait for a success message:\n\n'
    + 'http://ec2-13-57-188-159.us-west-1.compute.amazonaws.com/AmazonMWS/MarketplaceWebServiceProducts/Functions/CreateItemArray.php');
}

function palletMWS() {
  /**
  * This script uses the Amazon Products API in tandem with
  * http://ec2-13-57-188-159.us-west-1.compute.amazonaws.com PHP scripting to create a JSON file of the
  * items in the SCRAP sheet currently waiting to be listed.
  *
  * Only oversize items marked with a 'P' will be included.
  *
  * Use this function in tandem with the populateMWS() and
  * postListings() functions to list products on Amazon.
  */

  SpreadsheetApp.getUi().alert(
    'Go to the following URL and wait for a success message:\n\n'
    + 'http://ec2-13-57-188-159.us-west-1.compute.amazonaws.com/AmazonMWS/MarketplaceWebServiceProducts/Functions/CreatePalletArray.php');
}

function electronicsMWS() {
  /**
  * This script uses the Amazon Products API in tandem with
  * http://ec2-13-57-188-159.us-west-1.compute.amazonaws.com PHP scripting to create a JSON file of the
  * items in the SCRAP sheet currently waiting to be listed.
  *
  * Only small items marked with an 'E' will be included.
  *
  * Use this function in tandem with the populateMWS() and
  * postListings() functions to list products on Amazon.
  */

  SpreadsheetApp.getUi().alert(
    'Go to the following URL and wait for a success message:\n\n'
    + 'http://ec2-13-57-188-159.us-west-1.compute.amazonaws.com/AmazonMWS/MarketplaceWebServiceProducts/Functions/CreateElectronicArray.php');
}

function populateMWS() {
   /**
   * This script accomplishes the following tasks:
   *  1. Pull json file from MWS server
   *  2. Convert json into multidimensional array
   *  3. Push array into MWS tab.
   */

   // Fetch the json array from website and parse into JS object.
   var response = UrlFetchApp.fetch('http://ec2-13-57-188-159.us-west-1.compute.amazonaws.com/AmazonMWS/MarketplaceWebServiceProducts/Functions/MWS.json');
   var json = response.getContentText();
   var data = JSON.parse(json);

   // Convert data object into multidimensional array.
   // Ordering is same as in MWS tab.
   var itemCount = data.length;
   var itemArray = makeArray(12, itemCount, "");
   for (i = 0; i < itemCount; i++) {
     var item = data[i];
     try {
       itemArray[i] = ([
         item.SellerSKU,
         item.Title,
         item.UPC,
         item.ASIN,
         item.ItemPrice.toFixed(2),
         item.Dimensions.Weight,
         item.Dimensions.Length,
         item.Dimensions.Width,
         item.Dimensions.Height,
         item.Condition,
         item.Comment,
         ""
       ]);
     } catch(err) {
       itemArray[i] = ([
         item.SellerSKU,
         item.Title,
         item.UPC,
         item.ASIN,
         item.ItemPrice,
         item.Dimensions.Weight,
         item.Dimensions.Length,
         item.Dimensions.Width,
         item.Dimensions.Height,
         item.Condition,
         item.Comment,
         ""
       ]);
     }
   }

   // Push array into MWS tab.
   var sheetMWS = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('MWS');
   sheetMWS.getRange(2, 1, sheetMWS.getMaxRows()-1, 12).clearContent().setBackground('white');
   var range = sheetMWS.getRange(2, 1, itemCount, 12);
   range.setValues(itemArray);

   // Highlight undefined prices and wonky dimensions.
   var prices = sheetMWS.getRange(2, 5, itemCount, 6).getValues();
   for (i = 0; i < itemCount; i++) {
     if (prices[i][0] == "MANUAL" || prices[i][0] == "") {
       sheetMWS.getRange(2+i, 5).setBackground('red');
     }
     if (Number(prices[i][1]) < 1 || Number(prices[i][1]) > 100 || Number(prices[i][1]) == "NaN") {
       sheetMWS.getRange(2+i, 6).setBackground('red');
     }
     if (Number(prices[i][2]) < 1 || Number(prices[i][2]) > 75 || Number(prices[i][2]) == "NaN") {
       sheetMWS.getRange(2+i, 7).setBackground('red');
     }
     if (Number(prices[i][3]) < 1 || Number(prices[i][3]) > 75 || Number(prices[i][3]) == "NaN") {
       sheetMWS.getRange(2+i, 8).setBackground('red');
     }
     if (Number(prices[i][4]) < 1 || Number(prices[i][4]) > 75 || Number(prices[i][4]) == "NaN") {
       sheetMWS.getRange(2+i, 9).setBackground('red');
     }
   }
 }

function postListings() {
  /**
  * This script uses the XML exporter web apps in combination
  * with http://ec2-13-57-188-159.us-west-1.compute.amazonaws.com PHP scripts to send product listings
  * to Amazon.
  *
  * Use this function in tandem with the auditListings() function
  * to verify completion of the listings.
  */

  // Send product feeds to Amazon.
  var response = UrlFetchApp.fetch('http://ec2-13-57-188-159.us-west-1.compute.amazonaws.com/AmazonMWS/MarketplaceWebService/Functions/CreateNewListings.php?pass=K1@$run');
}

function createShipments() {
  /**
  * This script uses the Amazon FBAInboundMWS API in tandem with
  * http://ec2-13-57-188-159.us-west-1.compute.amazonaws.com PHP scripting to create a shipment with all items
  * in the MWS sheet.
  */

  SpreadsheetApp.getUi().alert(
    'Go to the following URL and wait for a success message:\n\n'
    + 'http://ec2-13-57-188-159.us-west-1.compute.amazonaws.com/AmazonMWS/FBAInboundServiceMWS/Functions/MasterShipment.php');
}

function shipLTL() {
  /**
  * This script uses the Amazon FBAInboundMWS API in tandem with
  * http://ec2-13-57-188-159.us-west-1.compute.amazonaws.com PHP scripting to create a palleted shipment with
  * all items in the MWS sheet.
  */

  SpreadsheetApp.getUi().alert(
    'Go to the following URL and wait for a success message:\n\n'
    + 'http://ec2-13-57-188-159.us-west-1.compute.amazonaws.com/AmazonMWS/FBAInboundServiceMWS/Functions/PalletShip.php');
}

function shipElectronics() {
  /**
  * This script uses the Amazon FBAInboundMWS API in tandem with
  * http://ec2-13-57-188-159.us-west-1.compute.amazonaws.com PHP scripting to create a small item, all-in-one-box
  * shipment with all items in the MWS sheet.
  */

  SpreadsheetApp.getUi().alert(
    'Go to the following URL and wait for a success message:\n\n'
    + 'http://ec2-13-57-188-159.us-west-1.compute.amazonaws.com/AmazonMWS/FBAInboundServiceMWS/Functions/ElectronicShip.php');
}

function importShipments() {
// Fetch the json array from website and parse into JS object.
  var response = UrlFetchApp.fetch('http://ec2-13-57-188-159.us-west-1.compute.amazonaws.com/AmazonMWS/FBAInboundServiceMWS/Functions/shipID.json');
  var json = response.getContentText();
  var data = JSON.parse(json);

  // Convert data object into multidimensional array.
  var shipments = Object.keys(data);
  var shipCount = shipments.length;
  var itemArray = {};
  var itemCount = [];
  var k = 0;
  for (var i = 0; i < shipCount; i++) {
    var id = shipments[i];
    itemCount[i] = data[id].length;

    for (var j = 0; j < itemCount[i]; j++) {
      itemArray[data[id][j]] = id;
    }
  }

  // Initialize sheet variables.
  var sheetMWS = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('MWS');
  var lastRow = sheetMWS.getLastRow();
  var SKUs = sheetMWS.getRange(1, 1, lastRow).getValues();

  // Import shipmentId's into sheet.
  sheetMWS.getRange(2, 12, lastRow-1, 1).clearContent();
  for (var i = 1; i < lastRow; i++) {
    sheetMWS.getRange(i+1, 12).setValue(itemArray[SKUs[i][0]]);
  }
}

