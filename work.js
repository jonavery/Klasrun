function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Automation Menu')
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

function auditListings() {
  /*
  This script accomplishes the following tasks:
    1. Search the Listings sheet for missing information
      i.e. measurements, initials, AER designation
    2. Make easy fixes if possible (AER designation)
    3. Highlight problem entries blue.
    4. Make relevant notes in REASON column.
    5. Move audit population to top of sheet.
  */
  
  // Initializing Work Listings sheet. Id is used rather than Active Spreadsheet to accommodate potential relocation.
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
