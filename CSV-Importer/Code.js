function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('CSV Importer')
      .addItem('Import CSV', 'showDialog')
      .addToUi();
}

function showDialog() {
  var htmlOutput = HtmlService.createHtmlOutputFromFile('CSVImportDialog.html')
      .setWidth(400)
      .setHeight(450);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'CSV Importer');
}

function importCSV(data, selectedColumns, importOption, filterText) {
  Logger.log("Importing CSV data: " + data);

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();
  
  var csvData = Utilities.parseCsv(data);
  // Filter data based on selected columns and filter text
  csvData = filterData(csvData, selectedColumns, filterText);

  if (importOption === "append") {
    // Append data as new columns
    for (var i = 0; i < selectedColumns.length; i++) {
      var columnIndex = selectedColumns[i];
      var columnData = [];
      
      for (var j = 0; j < csvData.length; j++) {
        if (csvData[j][columnIndex]) {
          columnData.push([csvData[j][columnIndex]]);
        } else {
          columnData.push(['']);
        }
      }
  
      Logger.log("Importing column " + columnIndex + ": " + columnData);
  
      sheet.getRange(1, sheet.getLastColumn() + 1, columnData.length, 1).setValues(columnData);
    }
  } else if (importOption === "clear") {
    // Clear data in the current sheet and paste new data
    clearCurrentSheet(spreadsheet);
    var numRows = csvData.length;
    var numCols = selectedColumns.length;

    // Set the range to paste data starting from cell A1
    var range = sheet.getRange(1, 1, numRows, numCols);

    // Set the values in the range
    range.setValues(csvData);
  } else if (importOption === "new-sheet-option") {
    // Create a new sheet and paste data into it
    createNewSheet(spreadsheet, csvData, selectedColumns);
    createNewSheet(spreadsheet, data, selectedColumns);

  }
}

function filterData(data, selectedColumns, filterText) {
  return data.filter(function (row) {
    return selectedColumns.some(function (colIndex) {
      var cellValue = row[colIndex].toLowerCase(); // Convert to lowercase for case-insensitive matching
      return cellValue.includes(filterText.toLowerCase());
    });
  });
}

function createNewSheet(spreadsheet, data, selectedColumns, filterText) {
  // Create a new spreadsheet
  var newSpreadsheet = SpreadsheetApp.create('New CSV Import');

  // Open the newly created spreadsheet
  var newSheet = newSpreadsheet.insertSheet('Imported Data');
  var csvData = Utilities.parseCsv(data);
  
  // Filter data based on selected columns and filter text
  csvData = filterData(csvData, selectedColumns, filterText);

  var numRows = csvData.length;
  var numCols = selectedColumns.length;

  // Set the range to paste data starting from cell A1
  var range = newSheet.getRange(1, 1, numRows, numCols);

  // Set the values in the range
  range.setValues(csvData.map(function (row) {
    return selectedColumns.map(function (colIndex) {
      return row[colIndex];
    });
  }));
}

function clearCurrentSheet(spreadsheet) {
  var sheet = spreadsheet.getActiveSheet();
  sheet.clear();
}

function getCSVColumns(data) {
  var csvData = Utilities.parseCsv(data);
  if (csvData.length > 0) {
    return csvData[0];
  } else {
    return [];
  }
}

function filterData(data, selectedColumns, filterText) {
  return data.filter(function (row) {
    return selectedColumns.some(function (colIndex) {
      var cellValue = row[colIndex].toLowerCase(); // Convert to lowercase for case-insensitive matching
      return cellValue.includes(filterText.toLowerCase());
    });
  });
}
