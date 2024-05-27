function fetchDataForSEO() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var settingsSheet = ss.getSheetByName("DataForSEO Settings");
  var outputSheet = ss.getSheetByName("DataForSEO Output");

  // Get settings from settings sheet
  var range = settingsSheet.getRange("A6:B" + settingsSheet.getLastRow());
  var values = range.getValues();
  var countryName = settingsSheet.getRange("B2").getValue();
  var languageName = settingsSheet.getRange("B3").getValue();
  var countryCode = settingsSheet.getRange("C2").getValue();
  var languageCode = settingsSheet.getRange("C3").getValue();
  var device = settingsSheet.getRange("B4").getValue();

  // Write headers to output sheet if it's empty
  if (outputSheet.getLastRow() == 0) {
    outputSheet.appendRow(["Keyword", "SERP Country", "SERP Language", "Search Device", "SERP Items"]);
  }

  // Iterate through each query
  for (var i = 0; i < values.length; i++) {
    var query = values[i][0];
    var checkbox = values[i][1];

    if (query && !checkbox) {
      var responseData = fetchSERPData(query, countryCode, languageCode, countryName, languageName, device);

      // Write data to output sheet
      outputSheet.appendRow([query, countryName, languageName, device, responseData.join(",")]);
    }
  }
}

function fetchSERPData(query, countryCode, languageCode, countryName, languageName, device) {
  var apiUrl = 'https://api.dataforseo.com/v3/serp/google/organic/live/advanced';
  var username = ''; // Replace with your DataForSEO username
  var password = ''; // Replace with your DataForSEO password

  var payload = {
    "keyword": encodeURI(query),
    "language_code": languageCode,
    "location_code": countryCode,
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
