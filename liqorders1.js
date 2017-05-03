function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Automation Menu')
    .addItem('Update VLOOKUPs', 'priceSearch')
    .addToUi();
}

function priceSearch() {
  /**
  * This script accomplishes the following:
  *   1. Set spacing between first and second order to 3 blank rows
  *   2. Copy formulas from Subtotal and My Buy Price
  *   3. Move UPC and Notes information
  *   4. Find and print sum of Count column
  *   5. Update VLOOKUP range
  *   6. Copy A/E/R and Amazon formulas
  *   7. Cache and paste values to overwrite the formulas
  **/
  
  // Initialize sheet.
  var sheet = SpreadsheetApp.getActiveSheet();
  
  // Set spacing between first and second order to 3 blank rows.
  
  // Copy formulas from Subtotal and My Buy Price.
  
  // Move UPC and Notes information.
  
  // Find and print sum of Count column.
  
  // Update VLOOKUP range.
  
  // Copy A/E/R and Amazon formulas.
  
  // Cache and paste values to overwrite the formulas.
  
}
