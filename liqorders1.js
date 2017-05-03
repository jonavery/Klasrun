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
  var valuesSubtotal = sheet.getRange("A3:H4").getValues();
  sheet.getRange(lastItemRow+1, 1, 2, 8).setValues(valuesSubtotal);
  
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
  var upc = sheet.getRange(6, 6, itemCount).getValues();
  var notes = sheet.getRange(7, 7, itemCount-1).getValues();
  sheet.getRange(6, 4, itemCount).clear().setValues(upc);
  sheet.getRange(7, 1, itemCount-1).setValues(notes);
  sheet.getRange(6, 5, itemCount).clear();
  
  // Find and print sum of Count column.
  var countA1 = sheet.getRange(6, 3, itemCount).getA1Notation();
  sheet.getRange(lastItemRow+1, 3).setFormula("=SUM("+countA1+")");
  
  // Update VLOOKUP range.
  var vlookupA1 = sheet.getRange(lastItemRow+6, 2, sheet.getLastRow()-lastItemRow-5, sheet.getLastColumn()-1).getA1Notation();
  sheet.getRange(2, 5).setFormula("=VLOOKUP($B2,"+vlookupA1+",4,0)");
  sheet.getRange(2, 6).setFormula("=VLOOKUP($B2,"+vlookupA1+",5,0)");
  sheet.getRange(2, 7).setFormula("=VLOOKUP($B2,"+vlookupA1+",6,0)");
  
  // Copy A/E/R and Amazon formulas.
  var formulas = sheet.getRange(2, 5, 1, 3).getValues();
  for (i=0; i < itemCount; i++) {
    for (j=0; j < formulas[0].length; j++) {
      sheet.getRange(6+i, 5+j).setValue(formulas[0][j]);
    }
  }
  
  // Cache and paste values to overwrite the formulas.
  
}
