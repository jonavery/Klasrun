function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Automation Menu')
    .addItem('Update Item in Liquidation', 'updateAssorted')  
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

function updateAssorted() {
  /**
  * This script gets an SKU from the user and updates the item in
  * Liquidation to match it.
  */
  
  // Cache spreadsheet ID's
  var liquidID = "1Vfzm-MogYyJB88YFmA4oWQ1M4bpT4Onw-ZGeFSpqseo";
  var workID = "1w28MV69JaR99e2m-2hveMLDY_Ukbz1Meg-9RQK0-sik";
  
  // Initialize Work and Liquidation sheets.
  var sheetListings = SpreadsheetApp.openById(workID).getSheetByName("Listings");
  var sheetLiquid = SpreadsheetApp.openById(liquidID).getSheetByName("Liquidation Orders");
  
  // Prompt user for SKU.
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Enter the SKU:');
  
  if (response.getSelectedButton() == ui.Button.OK) {
    // Cache user-given SKU and initialize work and liquidation SKU's.
    var sku = parseInt(response.getResponseText());
    var workSKU = getCol(sheetListings.getRange(1, 2, sheetListings.getLastRow()).getValues(), 0);
    var liquidSKU = getCol(sheetLiquid.getRange(1, 1, sheetListings.getLastRow()).getValues(), 0);
    
    // Find index of SKU in work and liquidation.
    var workIndex = workSKU.indexOf(sku);
    var liquidIndex = liquidSKU.indexOf(sku);
    Logger.log("Work: " + workIndex)
    Logger.log("Liquid: " + liquidIndex)
    
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
    SpreadsheetApp.getUi().alert(
      'Item title updated from "' + title + '" to "' + workValues[0][0] + '".');
    
  } else {
    SpreadsheetApp.getUi().alert('No changes made.');
  }
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
  
  // Initialize Work Listings sheet. Id is used rather than Active Spreadsheet to accommodate potential relocation.
  var sheetListings = SpreadsheetApp.openById("1w28MV69JaR99e2m-2hveMLDY_Ukbz1Meg-9RQK0-sik").getSheetByName("Listings");
  var allListings = sheetListings.getDataRange().getValues();
  
  // Loop through each row and cache tested item rows.
  var doneListings = [];
  // var today = new Date().getDate();
  var today = "14";
  for (i=0; i < allListings.length; i++) {
    var aerMeasurements = allListings[i].slice(6).join("");
    if (aerMeasurements != "") {
      doneListings.push(allListings[i]);
    }
  }
//  for (i=0; i < doneListings.length(); i++) {
//    if (!doneListings[i][0].includes(today)) {
//        doneListings.splice(i, 1);
//    }
//  }
  Logger.log("Rows: " + doneListings.length);
  Logger.log("Cols: " + doneListings[0].length);
}
