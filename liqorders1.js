function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Automation Menu')
    .addItem('Liquidation Price Lookup', 'liqPriceSearch')
    .addItem('Blackwrap Price Lookup', 'blackPriceSearch')
    .addToUi();
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
  
  // Make items returns if belonging to certain brands.
  var banned = ['Gourmia', 'Cheftronic', 'Oliso', 'Wondermill', 'SKG', 'KitchenAid'];
  var items = sheet.getRange(6, 2, itemCount).getValues();
  for (i=0; i < itemCount; i++) {
    for (j=0; j < banned.length; j++) {
      if (items[i][0].indexOf(banned[j]) != -1) {
        sheet.getRange(6+i, 5).setValue('R');
      }
    }
  }
  
  
  // Cache and paste values to overwrite the formulas.
  var vlookupValues = sheet.getRange(6, 5, itemCount, 3).getValues();
  sheet.getRange(6, 5, itemCount, 3).setValues(vlookupValues);
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
  
  // Make items returns if belonging to certain brands.
  var banned = ['Gourmia', 'Cheftronic', 'Oliso', 'Wondermill', 'SKG', 'KitchenAid'];
  var items = sheet.getRange(6, 2, itemCount).getValues();
  for (i=0; i < itemCount; i++) {
    for (j=0; j < banned.length; j++) {
      if (items[i][0].indexOf(banned[j]) != -1) {
        sheet.getRange(6+i, 5).setValue('R');
      }
    }
  }
  
  
  // Cache and paste values to overwrite the formulas.
  var vlookupValues = sheet.getRange(6, 5, itemCount, 3).getValues();
  sheet.getRange(6, 5, itemCount, 3).setValues(vlookupValues);
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
  var data = JSON.parse(json);
  
  // Convert data object into multidimensional array.
  // Ordering is same as in MWS tab.
  var itemCount = data.length;
  var itemArray = makeArray(11, itemCount, "");
  for (i = 0; i < itemCount; i++) {
    var item = data[i];
    itemArray[i] = ([
      // item.Index,
      item.Title,
      item.UPC,
      item.ASIN
      // item.Price
    ]);
  }

  // Push array into MWS tab.
  var sheet6 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet6');
  var range = sheetMWS.getRange(2, 1, itemCount, 11);
  range.setValues(itemArray);
  
  // Highlight undefined entries that will not be listed.
  var prices = sheetMWS.getRange(2, 5, itemCount).getValues();
  for (i = 0; i < itemCount; i++) {
    if (prices[i][0] == "undefined") {
      sheetMWS.getRange(2+i, 1, 1, 11).setBackground('red');
    }
  }
}
