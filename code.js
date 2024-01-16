function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('MENU UI')
    .addItem('Add Entry', 'enterNewEntrydrop')
    .addItem('Search Using Filters ', 'ReadSearch')
    .addItem('Edit Entry', 'editEntry')
    .addItem('Delete Entry', 'deleteRow')
    .addToUi();
}

function enterNewEntrydrop() {
  try {
    // Get the active spreadsheet and sheet
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getActiveSheet();

    // Get column headers from the first row
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    // Create an HTML form for headers
    var htmlOutput = HtmlService.createHtmlOutputFromFile('EntryForm')
      .setWidth(500)
      .setHeight(500);

    // Pass headers to the HTML form
    htmlOutput.append('<form><input type="hidden" name="headers" value="' + headers.join(',') + '">');
    for (var i = 0; i < headers.length; i++) {
      htmlOutput.append('<label for="header' + i + '">' + headers[i] + ' : </label>');
           
      if (headers[i] === 'Date of Birth') {
        htmlOutput.append('<input type="date" id="header' + i + '" name="header' + i + '">');
         htmlOutput.append('<br><br>');
      }
      else if (headers[i] === 'Favorite Subject') { // If the header is 'DropdownHeader', add a dropdown
        htmlOutput.append('<select id="header' + i + '" name="header' + i + '">');
        htmlOutput.append('<option value="English">English</option>');
        htmlOutput.append('<option value="Maths">Maths</option>');
        htmlOutput.append('<option value="Science">Science</option>');
        htmlOutput.append('</select><br><br>');
      } 
      else {
        htmlOutput.append('<input type="text" id="header' + i + '" name="header' + i + '"><br><br>');
      }
    }
    htmlOutput.append('<input type="button" class="subbtn" value="Submit" onclick="submitForm()"></form>');

    // Display the form as a dialog
    var formResponse = SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Enter New Entry');
  } catch (error) {
    Logger.log("Error adding new entry: " + error.message);
    SpreadsheetApp.getUi().alert("Error: " + error.message);
  }
}

function appendRow(data) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.appendRow(data);
  SpreadsheetApp.getUi().alert('Data has been stored successfully!');
  return true;
}


function editEntry(){
  var ui = SpreadsheetApp.getUi();
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet=spreadsheet.getActiveSheet();

  var headers=sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  
  var htmlOutput = HtmlService.createHtmlOutputFromFile('EntryForm')
      .setWidth(600)
      .setHeight(500);

  
  var primarykey = ui.prompt('Enter the Student ID of the student to edit:', ui.ButtonSet.OK_CANCEL).getResponseText();

  // Find the row with the specified roll number
  var range = sheet.getRange('A:A');
  var values = range.getValues();
  var rowIndex = -1;
  var l=getFilledRowCount()
  for (var i = 0; i < l; i++) {
    if (values[i][0] == primarykey) {
      rowIndex = i + 1; // Adding 1 to convert from zero-based index to 1-based index
      break;
    }
  }
  if (rowIndex!==-1){
  rowData = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];

  htmlOutput.append('<form><input type="hidden" name="headers" value="' + headers.join(',') + '">');
  for (var i = 0; i < headers.length; i++){
        htmlOutput.append('<div class="form-group">');
        htmlOutput.append('<label for="' + headers[i] + '">' + headers[i] + ' : </label>');

/*
          if (headers[i] === 'Date of Birth')
          
        {htmlOutput.append('<input type="date" id="' + headers[i] + '" name="' + headers[i] + '" value="' + rowData[i] + '">');
        htmlOutput.append('</div> <br></br>');}
          
      else 
     */ 
      if (headers[i] === 'Favorite Subject') { // If the header is 'DropdownHeader', add a dropdown
        htmlOutput.append('<select id="' + headers[i] + '" name="'+ headers [i]+ '"value="' + rowData[i] + '">');
        htmlOutput.append('<option value="English">English</option>');
        htmlOutput.append('<option value="Maths">Maths</option>');
        htmlOutput.append('<option value="Science">Science</option>');
        htmlOutput.append('</select><br><br>');
      } 
else{

        {htmlOutput.append('<input type="text" id="' + headers[i] + '" name="' + headers[i] + '" value="' + rowData[i] + '">');
        htmlOutput.append('</div> ');}
  }
  }


  htmlOutput.append('<input type="button"  class="subbtn" value="Submit" onclick="editedForm('+rowIndex+')"></form>');

//htmlOutput.append('<input type="button" value="Submit" onclick="editedForm(' + rowIndex + ')"></form>');
   ui.showModelessDialog(HtmlService.createHtmlOutput(htmlOutput).setWidth(700).setHeight(500), 'Edit Row Data');
    } else {
      ui.alert('Row not found with the specified primary key.');
    }}
  

 function updateRowData(datas,rowIndex) {
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getActiveSheet();
    var headers = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
    var newRowData = [];   
   
    for (var i = 0; i < headers.length; i++) {
      newRowData.push(headers[i]);
    }
 // Loop through each value in the row and update if a new value is provided
  for (var i = 0; i < datas.length; i++) {
    if (datas[i] !== undefined) {
      newRowData[i] = datas[i];
    }
  }

    sheet.getRange(rowIndex, 1, 1, newRowData.length).setValues([newRowData]);
 //   var formResponse = SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Enter New Entry');

    // Indicate success
    SpreadsheetApp.getUi().alert('Data has been updated successfully!');
    return 'Row updated successfully!';

  } catch (error) {
    // Handle errors gracefully
    Logger.log(error);
    return 'An error occurred while updating the row. Please try again.';
  }
}


function closeDialog(message) {
  SpreadsheetApp.getUi().getActiveDialog().hide();
  SpreadsheetApp.getUi().alert(message);
}

function deleteRow() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var ui = SpreadsheetApp.getUi();

  // Prompt user to enter the ID with callback
  var response = ui.prompt('Delete row' , 'Enter the Student ID of the row to delete', ui.ButtonSet.YES_NO) 
    if (response.getResponseText()) {
      var uniqueId = parseInt(response.getResponseText());
      
      // Find the row with the matching unique ID
      var rowIndex = findRow(uniqueId);
      if (rowIndex > 0) {
        // Confirm deletion with another prompt
        var react = ui.alert('Are you sure you want to delete row ' + rowIndex + '?', ui.ButtonSet.YES_NO);
          if (react === ui.Button.YES) {
            sheet.deleteRow(rowIndex);
            ui.alert('Row deleted successfully!');
          }
        }   else {
            ui.alert('ID not found!')
        } 
      } else {
        ui.alert('ID not found!');
      }
    }

function findRow(uniqueId) {
  // Loop through all rows in the sheet
  for (var i = 1; i <= SpreadsheetApp.getActiveSheet().getLastRow(); i++) {
    // Check if the value in the specified column matches the unique ID
    if (SpreadsheetApp.getActiveSheet().getRange(i,1).getValue() === uniqueId) {
      return i; // Return the row index if found
    }
  }
  return 0; // Return 0 if not found
}

  function getFilledRowCount() {
  // Assuming your data is in the first column (column A)
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();

  // Iterate backward to find the last non-empty row
  for (var row = lastRow; row > 0; row--) {
    var value = sheet.getRange(row, 1).getValue(); // Assuming data is in column A
    if (value !== "") {
      // The first non-empty cell in column A is found
      var filledRowCount = row;
      Logger.log("Number of filled rows: " + filledRowCount);
      return filledRowCount;
    }
  }
  // If the loop completes without finding any non-empty cell
  Logger.log("No filled rows found.");
  return 0;
}
 

//CRUD : Read operation 
//Search with filters using drop downs


function ReadSearch() {
  var htmlOutput = HtmlService.createHtmlOutputFromFile('Page')
      .setWidth(400)
      .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Select Filters For Search');

 // return htmlOutput;
// htmlOutput.append('<script>window.onload = function() { PageSearch(); }</script>');
}

function getFilterHeaders() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return headers.filter(function(header) {
    return header !== ""; // Filter out empty headers
  });
}

function getUniqueKeys(filterHeader) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var headerRow = data[0];
  var headerIndex = headerRow.indexOf(filterHeader);

  if (headerIndex !== -1) {
    var uniqueKeys = [];
    for (var i = 1; i < data.length; i++) {
      var key = data[i][headerIndex];
      if (uniqueKeys.indexOf(key) === -1 && key !== "") {
        uniqueKeys.push(key);
      }
    }

    // Debugging statements
    console.log('Original uniqueKeys:', uniqueKeys);

    // If the header is a date, format the unique keys
    if (Object.prototype.toString.call(uniqueKeys[0]) === '[object Date]') {
      uniqueKeys = uniqueKeys.map(function(date) {
        var formattedDate = Utilities.formatDate(date, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 'dd/MM/yyyy');
        console.log('Formatted date:', formattedDate);
        return formattedDate;
      });
    }

    console.log('Final uniqueKeys:', uniqueKeys);

    return uniqueKeys;
  } else {
    return [];
  }
}

function getData(filterHeader, uniqueKey) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var headerRow = data[0];
  var headerIndex = headerRow.indexOf(filterHeader);

  if (headerIndex !== -1) {
    var result = [headerRow];
    for (var i = 1; i < data.length; i++) {
      if (data[i][headerIndex] == uniqueKey) {
        result.push(data[i]);
      }
    }

    if (result.length > 1) {
      showTable(result);
    } else {
      SpreadsheetApp.getUi().alert('Selected unique key not found.');
    }
  } else {
    SpreadsheetApp.getUi().alert('Selected filter header not found.');
  }
}


function showTable(data) {
  if (data && Array.isArray(data) && data.length > 1) {
    var htmlContent = '<table border="1">';

    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      htmlContent += '<tr>';
      
      for (var j = 0; j < row.length; j++) {
        var cell = row[j];
        // Format date cells to DD/MM/YYYY
        if (Object.prototype.toString.call(cell) === '[object Date]') {
          cell = Utilities.formatDate(cell, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 'dd/MM/yyyy');
        }
        htmlContent += '<td>' + cell + '</td>';
      }

      htmlContent += '</tr>';
    }

    htmlContent += '</table>';

    var htmlOutput = HtmlService.createHtmlOutput(htmlContent)
        .setWidth(500)
        .setHeight(300);

    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Search Results ');
  } else {
    SpreadsheetApp.getUi().alert('No data to display.');
  }
}
