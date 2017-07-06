function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Automation Menu')
    .addItem('Update Item in Liquidation By SKU', 'updateBySKU')
    .addItem('Update All Work Items in Liquidation', 'bulkUpdateLiquid')
    .addSeparator()
    .addItem('Highlight Future Listings by A/E/R', 'highlightAER')
    .addSeparator()
    .addItem('Populate MWS Tab', 'importPrices')
    .addItem('Post Listings', 'postListings')
    .addItem('Cancel Listings', 'cancelListings')
    .addItem('Audit Listings', 'auditListings')
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

function highlightAER() {
  /**
  * This script highlights each A/E/R cell according to its designation.
  * The script is coded to leave highlighted cells/rows alone.
  */
  
  // Initialize sheet and save values.
  var sheetWork = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Future Listing');
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
          activeRange.setBackground('white');
          break;
        case 'e':
        case 'E':
          activeRange.setBackground('#ff00ff');
          break;
        case 'r':
        case 'R':
          activeRange.setBackground('orange');
          break;
        case 'd':
        case 'D':
          activeRange.setBackground('red');
          break;
        case 'RHD':
          activeRange.setBackground('blue');
          activeRange.setFontColor('white');
          break;
        case 'c':
        case 'C':
          activeRange.setBackground('#cc3300');
          activeRange.setFontColor('white');
          break;
        default:
          activeRange.setBackground('gray');
          break;
      }
    }
  }
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
  var sheetListings = SpreadsheetApp.getActiveSheet();
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
  */
  
  // Cache spreadsheet ID's
  var liquidID = "1Xqsc6Qe_hxrWN8wRd_vgdBdrCtJXVlvVC9w53XJ0BUM";
  var workID = "1okDFF9236lGc4vU6W7HOD8D-3ak8e_zntehvFatYxnI";
  
  // Initialize Work and Liquidation sheets.
  var sheetListings = SpreadsheetApp.getActiveSheet();
  var sheetLiquid = SpreadsheetApp.openById(liquidID).getSheetByName("Liquidation Orders");
  
  // Prompt user for SKU.
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('This will update the liquidation sheet to match all items in the currently selected sheet.'
                         +'\n\nHit OK to continue.\nClose this window to cancel.');
  
  if (response == ui.Button.OK) {
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
    
    for (var i = 1; i < workSKU.length; i++) {
      // Find index of SKU in work and liquidation.
      var sku = parseInt(workSKU[i]);
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
    // Show changes to user.
    ui.alert(
      'Items updated: ' + updated
      + '\nItems already up to date: ' + notUpdated
      + '\nItems created: ' + created);
  }
}

function postListings() {
  /**
  * This script uses the XML exporter web apps in combination
  * with klasrun.com PHP scripts to send product listings
  * to Amazon.
  *
  * Use this function in tandem with the auditListings() function
  * to verify completion of the listings.
  */
  
  // Send product feeds to Amazon.
  var response = UrlFetchApp.fetch('http://klasrun.com/AmazonMWS/MarketplaceWebService/Functions/CreateNewListings.php');
}

function cancelListings() {
  /**
  * Run this script after postListings to cancel the feed submission.
  *
  * Only works before the feed has been processed by Amazon. Time is
  * of the essence here!
  */
  
  SpreadsheetApp.getUi().alert('ERROR: Script still in development.');
  
  // Call klasrun.com to cancel feed submission.
  
  
  // Report status of cancellation.
}

function auditListings() {
  /**
  * This script accomplishes the following tasks:
  *  1. Query amazon for listing status report.
  *  2. Report status of three listing feeds.
  *  3. When all three are complete, highlight items green in Scrap and MWS.
  *  4. Migrate completed listings from Scrap to Archive.
  */
  
  SpreadsheetApp.getUi().alert('ERROR: Script still in development.');
  
  // Grab status report from klasrun.com
  
  
  // Broadcast status report via email.
  
  
  // Initialize Scrap, MWS, and Archive sheets.
  
  
  // Highlight complete items green in Scrap and MWS sheets.
  
  
  // Move completed scrap entries into Archive sheet.
}
