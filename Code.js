function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Create a custom menu in Google Sheets
  ui.createMenu('CSV Importer')
    .addItem('Import CSV', 'showImportDialog')
    .addToUi();
}

function showImportDialog() {
  var htmlOutput = HtmlService.createHtmlOutputFromFile('Index')
    .setWidth(400)
    .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'CSV Import');
}

function importCSVData(csvData) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var csvArray = Utilities.parseCsv(csvData, ';'); // Use semicolon (;) as the delimiter

  // Clear the existing content in the sheet (optional)
  sheet.clear();
  
  // Get the starting cell where data will be inserted
  var startRow = 1; // Start from the first row
  var startCol = 1; // Start from the first column

  // Set the range for data insertion
  var numRows = csvArray.length;
  var numCols = csvArray[0].length;
  var range = sheet.getRange(startRow, startCol, numRows, numCols);
  
  // Insert the data
  range.setValues(csvArray);
}
