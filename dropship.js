function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Automation Menu')
    .addItem('Import VA Listings', 'importListings')
    .addToUi()
}

function importListings() {
    /**
    * This script updates values in the Ebay Dropshipping sheet
    * by importing all new entries in the VA Dropshipping sheets.
    *
    * It runs automatically once every hour.
    */
    
    // Cache spreadsheet ID's
    var masterID = "16HHFAByXMihdMPc6unV53zUKHMBeB-K59BGhoMEs4pM";
    var vaIDs = [
      "11DOJOXsEOo6de1HE3WwvJl4ROaz2nbGhTBLBEhEkdAs", // Ali
      "1KoFhl9HwvEbL7SgajKg1dRV7yDLmfQtZGTQWotBsGHA", // Javed
      "1CeASzstJ__tEa-RBrGz2wDgLRIWCUUYW7zwLq7Z3TAw"  // Mary
    ];
    
    // Initialize data values from master sheet.
    var sheetMaster = SpreadsheetApp.openById(masterID).getSheetByName("Listings");
    var masterValues = sheetMaster.getDataRange().getValues();
    var masterLastRow = sheetMaster.getLastRow();
    var masterItemNums = getCol(masterValues, 0);

    // Initialize counting variable.
    var created = 0;
  
    for (var i = 0; i < vaIDs; i++) { 
        // Initialize data values from VA sheet.
        var sheetVA = SpreadsheetApp.openById(vaIDs[i]).getSheetByName("Listings");
        var vaValues = sheetVA.getDataRange().getValues();
        var vaLastRow = sheetVA.getLastRow();
        var vaItemNums = getCol(vaValues, 0);
          
        for (var j = 1; j < vaItemNums.length; j++) {
            var itemID = Number(vaItemNums[j]);
            // Skip entry if no item number.
            if (isNaN(itemID) || itemID == "") {continue;}
            // Find index of SKU in work and liquidation.
            var masterIndex = masterItemNums.indexOf(itemID);
            // Skip entry if already in master sheet.
            if (masterIndex != -1) {continue;}
              
            // Skip entry if title is blank or already up to date.
            if (vaValues[j][6] == "") {
                continue;
            }
              
            // Create row to import new entry and log its position.
            var r = String(masterLastRow + 1);
            sheetMaster.insertRowAfter(masterLastRow);
             
            // Import values from VA sheet.
            sheetMaster.getRange(r, 1, 1, 11).setValue(vaValues[j]);
             
            // Setup liquidation formulas for new entry.
            sheetMaster.getRange(r, 4).setFormula('=IF(ISNA(VLOOKUP(C'+r+',BanList!A:A,1,0)),IF(ISNA(VLOOKUP(C'+r+',BanList!B:B,1,0)),"UNSURE","OK"),"BAN")');

            // Increment variables accordingly.
            masterLastRow++;
            created++;
        }
    }
    Logger.log('Listings imported: ' + created);
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
