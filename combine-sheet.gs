function importAllCSVFromFolder() {
    var folderName = "login-csv-27062024"; // Name of the folder containing CSV files
    var folder = getFolderByName(folderName);
    
    if (folder) {
      var files = folder.getFilesByType(MimeType.CSV);
      var fileArray = [];
  
      // Collect all files into an array
      while (files.hasNext()) {
        fileArray.push(files.next());
      }
  
      // Sort the files by last modified date in ascending order
      fileArray.sort(function(a, b) {
        return a.getLastUpdated() - b.getLastUpdated();
      });
  
      var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
      for (var i = 0; i < fileArray.length; i++) {
        var file = fileArray[i];
        var fileContent = file.getBlob().getDataAsString();
        var csvData = parseCsvWithCustomDelimiter(fileContent, '`'); // Using backtick as delimiter
        
        // Check if csvData has rows and columns
        if (csvData.length === 0 || csvData[0].length === 0) {
          Logger.log("Empty or invalid CSV file: " + file.getName());
          continue;
        }
        
        // Determine the number of columns based on the maximum number of columns found in the CSV data
        var numCols = getMaxColumns(csvData);
  
        // Ensure each row has the correct number of columns
        csvData = csvData.map(function(row) {
          while (row.length < numCols) {
            row.push(""); // Add empty strings if there are missing columns
          }
          return row.slice(0, numCols); // Trim rows to have exactly `numCols` columns
        });
        
        var numRows = csvData.length;
  
        // Create a new sheet for each file
        var sheetName = file.getName().replace(/\.csv$/i, ''); // Remove the .csv extension for the sheet name
        var sheet = spreadsheet.getSheetByName(sheetName);
      
        if (sheet) {
          // If the sheet already exists, clear its content
          sheet.clear();
        } else {
          // If the sheet doesn't exist, create a new one
          sheet = spreadsheet.insertSheet(sheetName);
        }
  
        // Set values in the new sheet
        sheet.getRange(1, 1, numRows, numCols).setValues(csvData);
      }
    } else {
      Logger.log("Folder not found: " + folderName);
    }
  }
  
  function getFolderByName(name) {
    var folders = DriveApp.getFoldersByName(name);
    if (folders.hasNext()) {
      return folders.next();
    } else {
      return null;
    }
  }
  
  function parseCsvWithCustomDelimiter(csvString, delimiter) {
    var rows = csvString.split('\n');
    var result = rows.map(function(row) {
      return row.split(delimiter);
    });
    return result;
  }
  
  function getMaxColumns(csvData) {
    var maxColumns = 0;
    csvData.forEach(function(row) {
      if (row.length > maxColumns) {
        maxColumns = row.length;
      }
    });
    return maxColumns;
  }
  