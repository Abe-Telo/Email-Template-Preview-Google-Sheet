// Get active sheet
function getColIndexByName(colName) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var numColumns = sheet.getLastColumn();
  var row = sheet.getRange(1, 1, 1, numColumns).getValues();
  for (i in row[0]) {
    var name = row[0][i];
    if (name == colName) {
      return parseInt(i) + 1;
    }
  }
  return -1;
}

//Get Static Sheet 
function getColIndexByNameStatic(sheet, colName) {
  // Get the range of cells in the first row of the sheet
  var range = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  // Get the values of the cells in the range as a two-dimensional array
  var values = range.getValues();
  // Loop through the array to find the index of the column with the specific name
  for (var i = 0; i < values[0].length; i++) {
    if (values[0][i] == colName) {
      // Return the index of the column
      return i + 1;
    }
  }
  // Return -1 if the column was not found
  return -1;
}
  
