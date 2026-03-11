function aa() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var result = [];
  
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var newRow = [];
    
    for (var j = 0; j < row.length; j++) {
      var cellValue = row[j];
      // Example transformation: Convert to uppercase
      newRow.push(cellValue.toString().toUpperCase());
    }
    
    result.push(newRow);
  }
  
  // Output the transformed data to a new sheet
  var outputSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Output") || SpreadsheetApp.getActiveSpreadsheet().insertSheet("Output");
  outputSheet.clear(); // Clear existing content
  outputSheet.getRange(1, 1, result.length, result[0].length).setValues(result);
}
