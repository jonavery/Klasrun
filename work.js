function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Automation Menu')
    .addItem('Update Item in Liquidation', 'updateAssorted')
    .addItem('Update All Work Items in Liquidation', 'bulkUpdateLiquid')
    .addSeparator()
    .addItem('Highlight Future Listings by A/E/R', 'highlightAER')
    .addSeparator()
    .addItem('Populate MWS Tab', 'importPrices')
    .addItem('Post Listings', 'postListings')
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
  var aerValues = getCol(workValues, 4);
  
  // Loop through A/E/R column and color cells with a switch statement.
  for (i=1; i<aerValues.length; i++) {
    if (workColors[i][0] == "#ffffff") {
      var activeRange = sheetWork.getRange(i+1, 2, 1, 5);
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
        default:
          activeRange.setBackground('gray');
          break;
      }
    }
  }
}

function updateAssorted() {
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
    var workSKU = getCol(sheetListings.getRange(1, 2, sheetListings.getLastRow()).getValues(), 0);
    var liquidSKU = getCol(sheetLiquid.getRange(1, 1, sheetLiquid.getLastRow()).getValues(), 0);
    
    // Find index of SKU in work and liquidation.
    var workIndex = workSKU.indexOf(sku);
    var liquidIndex = liquidSKU.indexOf(sku);
    Logger.log("Work: " + workIndex)
    Logger.log("Liquid: " + liquidIndex)
    
    if (liquidIndex == -1) {
      ui.alert('SKU not found in Liquidation.');
      return;
    }
    // Cache existing liquidation information.
    var title = sheetLiquid.getRange(liquidIndex+1, 3).getValue();
    var upc = sheetLiquid.getRange(liquidIndex+1, 5).getValue();
    var aer = sheetLiquid.getRange(liquidIndex+1, 7).getValue();
      
    // Copy work information into liquidation.
    var workValues = sheetListings.getRange(workIndex+1, 3, 1, 3).getValues();
    sheetLiquid.getRange(liquidIndex+1, 3).setValue(workValues[0][0]);
    sheetLiquid.getRange(liquidIndex+1, 5).setValue(workValues[0][1]);
    sheetLiquid.getRange(liquidIndex+1, 7).setValue(workValues[0][2]);
      
    // Show changes to user.
    ui.alert(
      'Item title updated from "' + title + '" to "' + workValues[0][0] + '".');
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
    var workValues = sheetListings.getRange(1, 2, workLastRow, 5).getValues();
    var liqLastRow = sheetLiquid.getLastRow();
    var liquidSKU = getCol(sheetLiquid.getRange(1, 1, liqLastRow).getValues(), 0);
    var liquidTitles = getCol(sheetLiquid.getRange(1, 3, liqLastRow).getValues(), 0);
    
    // Initialize counting variables.
    var updated = 0;
    var notUpdated = 0;
    var created = 0;
    
    for (i = 1; i < workSKU.length; i++) {
      // Find index of SKU in work and liquidation.
      var sku = parseInt(workSKU[i]);
      var workIndex = workSKU.indexOf(sku);
      var liquidIndex = liquidSKU.indexOf(sku);
      
      Logger.log("Work: " + workValues[i][0] + "\nLiq: " + liquidTitles[liquidIndex] + "\n");
      // Check if title is blank or already up to date.
      if (workValues[i][0] == "" || workValues[i][1] == liquidTitles[liquidIndex]) {
        notUpdated++;
        continue;
      }
      
      if (liquidIndex == -1) {
        created++;
        liquidIndex = liqLastRow;
        sheetLiquid.getRange(liquidIndex+1, 1).setValue(workValues[i][0]);
        sheetLiquid.getRange(liquidIndex+1, 2).setValue(todayDate());
        sheetLiquid.getRange(liquidIndex+1, 4).setValue("1");
        sheetLiquid.getRange(liquidIndex+1, 6).setValue("LIQUIDATION");
        sheetLiquid.getRange(liquidIndex+1, 8).setValue(workValues[i][4]);
      }
      
      // Copy work information into liquidation.
      sheetLiquid.getRange(liquidIndex+1, 3).setValue(workValues[i][1]);
      sheetLiquid.getRange(liquidIndex+1, 5).setValue(workValues[i][2]);
      sheetLiquid.getRange(liquidIndex+1, 7).setValue(workValues[i][3]);
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

function importPrices() {
  /**
  * This script accomplishes the following tasks:
  *  1. Pull json file from MWS server
  *  2. Convert json into multidimensional array
  *  3. Push array into MWS tab.
  */
  
  // Fetch the json array from website and parse into JS object.
  var response = UrlFetchApp.fetch('http://klasrun.com/AmazonMWS/MarketplaceWebServiceProducts/Functions/test.json');
  var json = response.getContentText();
  var data = JSON.parse(json);
  
  // Convert data object into multidimensional array.
  // Ordering is same as in MWS tab.
  var itemCount = data.length;
  var itemArray = makeArray(11, itemCount, "");
  for (i = 0; i < itemCount; i++) {
    var item = data[i];
    itemArray[i] = ([
      item.SellerSKU,
      item.Title,
      item.UPC,
      item.ASIN,
      item.Price,
      item.Dimensions.Weight,
      item.Dimensions.Length,
      item.Dimensions.Width,
      item.Dimensions.Height,
      item.Condition,
      item.Comment
    ]);
  }

  // Push array into MWS tab.
  var sheetMWS = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('MWS');
  var range = sheetMWS.getRange(2, 1, itemCount, 11).clearContent();
  range.setValues(itemArray);
  
  // Highlight undefined entries that will not be listed.
  var prices = sheetMWS.getRange(2, 5, itemCount).getValues();
  for (i = 0; i < itemCount; i++) {
    if (prices[i][0] == "undefined") {
      sheetMWS.getRange(2+i, 1, 1, 11).setBackground('red');
    }
  }
}

function postListings() {
  /**
  * This script accomplishes the following tasks:
  *  1. Use XML converter scripts to convert product information
  *  2. Send product feeds to Amazon
  *  3. Check for feed status updates from Amazon
  *  4. Report results.
  */
  
  SpreadsheetApp.getUi().alert('ERROR: Script still in development.');
}


function auditListings() {
  /**
  * This script accomplishes the following tasks:
  *  1. Search the Listings sheet for missing information
  *    i.e. measurements, initials, AER designation
  *  2. Make easy fixes if possible (AER designation)
  *  3. Highlight problem entries blue.
  *  4. Make relevant notes in REASON column.
  *  5. Move audit population to top of sheet.
  */
  
  SpreadsheetApp.getUi().alert('ERROR: Script still in development.');
}
