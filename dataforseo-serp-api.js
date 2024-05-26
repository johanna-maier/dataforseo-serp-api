function fetchDataForSEO() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var settingsSheet = ss.getSheetByName("DataForSEO Settings");
  var outputSheet = ss.getSheetByName("DataForSEO Output");

  // Get settings from settings sheet
  var queries = settingsSheet.getRange("A6:A").getValues().flat().filter(String);
  var countryCode = settingsSheet.getRange("C2").getValue();
  var languageCode = settingsSheet.getRange("C3").getValue();
  var device = settingsSheet.getRange("C4").getValue();

  // Clear output sheet
  outputSheet.clear();

  // Write headers to output sheet
  outputSheet.appendRow(["Keyword", "SERP Country", "SERP Language", "Search Device", "SERP Items"]);

  // Iterate through each query
  queries.forEach(function (query) {
    var responseData = fetchSERPData(query, countryCode, languageCode, device);

    // Write data to output sheet
    outputSheet.appendRow([query, countryCode, languageCode, device, responseData.join(",")]);
  });
}

function fetchSERPData(query, countryCode, languageCode, device) {
  var apiUrl = 'https://api.dataforseo.com/v3/serp/google/organic/live/advanced';
  var username = 'YOUR_USERNAME'; // Replace with your DataForSEO username
  var password = 'YOUR_PASSWORD'; // Replace with your DataForSEO password

  var payload = {
    "keyword": encodeURI(query),
    "language_code": languageCode,
    "location_code": countryCode
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
