/**
 * Sends a call to the DHIS2 Analytics Web API and returns the data to a Google Sheet.
 * 
 * @return DHIS2 analytics data in columns and rows in a Google Sheet.
 * @customfunction
 * 
 * Written by Rica Zamora Duchateau (Clinton Health Access Initiative).
*/

//@OnlyCurrentDoc
//Adds a custom menu to the Google Sheet
function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Import data from DHIS2")
    .addItem("Import to Raw Data sheet", "importCSVFromUrl")
    .addToUi();
}

//Displays an alert as a Toast message
function displayToastAlert(message) {
  SpreadsheetApp.getActive().toast(message, "Message:"); 
}

//Generates a string of months for the current plus previous 2 years,
//to be inserted into the DHIS2 API call as the period dimension
//Written with support from AskCodi
function generateMonthlyDates() {
  const currentDate = new Date();
  const currentYear = currentDate.getFullYear();
  const currentMonth = currentDate.getMonth() + 1; // Add 1 to get the current month (January is 0)
  let dates = '';

  for (let year = currentYear - 2; year <= currentYear; year++) {
    for (let month = 1; month <= 12; month++) {
      if (year === currentYear && month > currentMonth) {
        break; // Stop generating dates if we've gone past the current month and year
      }
      const date = `${year}${month.toString().padStart(2, '0')}`;
      dates += date + ';';
    }
  }

  return dates.slice(0, -1); // Remove the last semicolon
}

const API_KEY = '<INSERT DHIS2 API KEY HERE>';
// set the params object
const params = {
    headers: {
      Authorization: "Basic " + API_KEY
    }
}

//Construct authenticated DHIS2 Analytics API call and return results as CSV to the current sheet
function importCSVFromUrl() {
  // Set basic authentication key for DHIS2 instance
  var dateParameter = generateMonthlyDates();
  const url = '<INSERT DHIS2 ANALYTICS API REQUEST URL WITHOUT PERIOD DIMENSION>' + '&dimension=pe:' + dateParameter
  var contents = Utilities.parseCsv(UrlFetchApp.fetch(url,params));
  var sheetName = writeDataToSheet(contents, 'Raw Data');
  displayToastAlert("Data from DHIS2 was successfully imported into " + sheetName + ".");
}

//Clear current sheet and fill with contents from DHIS2 API call
function writeDataToSheet(data,sheetName) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(sheetName);
  sheet.clear(); // Clear sheet from existing content
  sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  return sheet.getName();
}
