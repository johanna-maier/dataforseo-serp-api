function getSERPData() {
  const sheetSettings = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DataForSEO Settings');
  const sheetOutput = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DataForSEO Output');

  // Clear previous output
  sheetOutput.clear();

  // Set output headers
  sheetOutput.getRange('A1:F1').setValues([['Keyword', 'SERP Country', 'SERP Language', 'Domain', 'Search Device', 'SERP Items']]);

  // Get search settings
  const searchQueries = sheetSettings.getRange('A6:A').getValues().filter(row => row[0]);
  const serpCountry = sheetSettings.getRange('C2').getValue();
  const serpLanguage = sheetSettings.getRange('C3').getValue();
  const searchDevice = sheetSettings.getRange('C4').getValue();

  // API credentials
  const username = 'your_username';  // replace with your DataForSEO API username
  const password = 'your_password';  // replace with your DataForSEO API password

  // Iterate through search queries
  searchQueries.forEach((row, index) => {
    const keyword = encodeURI(row[0]);
    const apiUrl = 'https://api.dataforseo.com/v3/serp/google/organic/live/advanced';
    const options = {
      method: 'post',
      headers: {
        'Authorization': 'Basic ' + Utilities.base64Encode(username + ':' + password),
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify([{
        keyword: keyword,
        language_code: serpLanguage,
        location_code: serpCountry
      }])
    };

    try {
      const response = UrlFetchApp.fetch(apiUrl, options);
      const data = JSON.parse(response.getContentText());

      if (data.status_code === 20000) {
        const task = data.tasks[0];
        const result = task.result[0];
        const items = result.item_types.join(', ');

        sheetOutput.appendRow([
          result.keyword,
          serpCountry,
          serpLanguage,
          result.se_domain,
          searchDevice,
          items
        ]);
      }
    } catch (error) {
      Logger.log('Error: ' + error.message);
    }
  });
}
