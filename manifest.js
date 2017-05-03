function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Automation Menu')
    .addItem('Update VLOOKUPs', 'priceSearch')
    .addToUi();
}

function priceSearch() {
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
  var items = sheet.getRange(6, 2, 100).getValues();
  for (i=0; i<items.length; i++) {
    if (items[i][0] == "") {
      var lastItemRow = i + 5;
      var itemCount = i;
      break;
    }
  }
  Logger.log("Last Item Row: " + lastItemRow)
  Logger.log("Item Count: " + itemCount)
  
  // Copy formulas from Subtotal and My Buy Price.
  var subRange = sheet.getRange("A3:H4");
  subRange.copyTo(sheet.getRange(lastItemRow+1, 1, 2, 8));
  
  // Find and cache the number of blank rows between first and second order.
  var blankCount = 0;
  for (i=itemCount; i<50; i++) {
    if (items[i][0] == "") {
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
      sheet.insertRowsAfter(lastItemRow+2, 5-blankCount);
      break;
    case (blankCount == 5):
      // Do nothing.
      break;
    case (blankCount > 5):
      // Delete rows.
      sheet.deleteRows(lastItemRow+3, blankCount-5);
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
  sheet.getRange(7, 1, itemCount-1).setValues(notes);
  sheet.getRange(6, 5, itemCount).clear();
  
  // Find and print sum of Count and Amazon columns.
  var countA1 = sheet.getRange(6, 3, itemCount).getA1Notation();
  var amazonA1 = sheet.getRange(6, 6, itemCount).getA1Notation();
  var feesA1 = sheet.getRange(6, 7, itemCount).getA1Notation();
  sheet.getRange(lastItemRow+1, 3).setFormula("=SUM("+countA1+")");
  sheet.getRange(lastItemRow+1, 6).setFormula("=SUM("+amazonA1+")");
  sheet.getRange(lastItemRow+1, 7).setFormula("=SUM("+feesA1+")");
  
  // Update VLOOKUP range.
  var vlookupA1 = sheet.getRange(lastItemRow+6, 2, sheet.getLastRow()-lastItemRow-5, sheet.getLastColumn()-1).getA1Notation();
  sheet.getRange(2, 5).setFormula("=VLOOKUP($B2,"+vlookupA1+",4,0)");
  sheet.getRange(2, 6).setFormula("=VLOOKUP($B2,"+vlookupA1+",5,0)");
  sheet.getRange(2, 7).setFormula("=VLOOKUP($B2,"+vlookupA1+",6,0)");
  
  // Copy A/E/R and Amazon formulas.
  var formulaRange = sheet.getRange(2, 5, 1, 3);
  for (i=0; i < itemCount; i++) {
      formulaRange.copyTo(sheet.getRange(6+i, 5, 1, 3));
  }
  
  // Cache and paste values to overwrite the formulas.
  var vlookupValues = sheet.getRange(6, 5, itemCount, 3).getValues();
  sheet.getRange(6, 5, itemCount, 3).setValues(vlookupValues);
}
