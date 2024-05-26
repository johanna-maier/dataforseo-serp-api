function fetchDataForSEO() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var settingsSheet = ss.getSheetByName("DataForSEO Settings");
  var outputSheet = ss.getSheetByName("DataForSEO Output");

  // Get settings from settings sheet
  var queries = settingsSheet.getRange("A6:A").getValues().flat().filter(String);
  var countryName = settingsSheet.getRange("B2").getValue();
  var languageName = settingsSheet.getRange("B3").getValue();
  var device = settingsSheet.getRange("B4").getValue();

  // Write headers to output sheet if it's empty
  if (outputSheet.getLastRow() == 0) {
    outputSheet.appendRow(["Keyword", "SERP Country", "SERP Language", "Search Device", "SERP Items"]);
  }

  // Iterate through each query
  queries.forEach(function (query) {
    var responseData = fetchSERPData(query, countryName, languageName, device);

    // Write data to output sheet
    outputSheet.appendRow([query, countryName, languageName, device, responseData.join(",")]);
  });
}

function fetchSERPData(query, countryName, languageName, device) {
  var apiUrl = 'https://api.dataforseo.com/v3/serp/google/organic/live/advanced';
  var username = ''; // Replace with your DataForSEO username
  var password = ''; // Replace with your DataForSEO password

  var payload = {
    "keyword": encodeURI(query),
    "language_name": languageName,
    "location_name": countryName,
    "device": device // Add the device information to the payload
  };

  var options = {
    method: 'post',
    headers: {
      Authorization: 'Basic ' + Utilities.base64Encode(username + ':' + password),
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify([payload])
  };

  var response = UrlFetchApp.fetch(apiUrl, options);
  var responseData = [];

  if (response.getResponseCode() == 200) {
    var data = JSON.parse(response.getContentText());
    if (data && data.tasks && data.tasks.length > 0 && data.tasks[0].result && data.tasks[0].result.length > 0) {
      responseData = data.tasks[0].result[0].item_types || [];
    }
  }

  return responseData;
}
